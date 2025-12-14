import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import chromium from "@sparticuz/chromium";
import puppeteer from "puppeteer-core";
import mammoth from "mammoth";

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: "50mb" }));

app.get("/", (req, res) => {
    res.send("HTML to PDF service is running. Use POST /convert with HTML content, or POST /convert-word-to-html with Word document.");
  });

// New endpoint for Word to HTML conversion using Hybrid approach:
// 1. Use mammoth.js to preserve merge fields ({{FieldName}})
// 2. Use docx-preview (via Puppeteer) to get perfect styling, colors, and design
// 3. Merge them: Apply docx-preview styling to mammoth content with merge fields
// This matches the approach used in the LWC component
app.post("/convert-word-to-html", async (req, res) => {
  let browser = null;
  try {
    const { file } = req.body; // Base64 encoded Word document
    const format = req.query.format?.toLowerCase() === "json" ? "json" : "html";

    if (!file) {
      return res.status(400).json({ success: false, error: "Word document file (base64) is required" });
    }

    // Decode base64 to buffer
    let wordBuffer;
    try {
      wordBuffer = Buffer.from(file, "base64");
    } catch (decodeError) {
      return res.status(400).json({ success: false, error: "Invalid base64 encoding: " + decodeError.message });
    }

    if (wordBuffer.length === 0) {
      return res.status(400).json({ success: false, error: "Word document is empty" });
    }

    console.log("Converting Word to HTML using Hybrid approach (mammoth.js for merge fields + docx-preview for styling). File size:", wordBuffer.length, "bytes");

    // Step 1: Use mammoth.js to extract HTML with merge fields preserved
    console.log("Step 1: Using mammoth.js to extract HTML with merge fields...");
    const styleMap = [
      "p[style-name='Title'] => h1.title:fresh",
      "p[style-name='Heading 1'] => h1:fresh",
      "p[style-name='Heading 2'] => h2:fresh",
      "p[style-name='Heading 3'] => h3:fresh",
      "r[style-name='Strong'] => strong",
      "r[style-name='Emphasis'] => em",
      "p[style-name='Normal'] => p"
    ];

    const mammothResult = await mammoth.convertToHtml(
      { buffer: wordBuffer },
      {
        styleMap: styleMap,
        includeDefaultStyleMap: true,
        includeEmbeddedStyleMap: true
      }
    );

    const mammothHtml = mammothResult.value;
    const messages = mammothResult.messages;
    
    if (messages && messages.length > 0) {
      console.log("Mammoth conversion messages:", messages);
    }

    if (!mammothHtml || mammothHtml.trim().length === 0) {
      return res.status(500).json({ 
        success: false, 
        error: "Mammoth conversion produced empty HTML content" 
      });
    }

    // Check if merge fields are preserved
    const mergeFieldPattern = /\{\{[^}]+\}\}/g;
    const mergeFields = mammothHtml.match(mergeFieldPattern) || [];
    console.log("Merge fields preserved in mammoth HTML:", mergeFields.length, "fields");
    if (mergeFields.length > 0) {
      console.log("✓ Sample merge fields from mammoth:", mergeFields.slice(0, 5));
      console.log("Mammoth HTML sample (first 1000 chars):", mammothHtml.substring(0, 1000));
    } else {
      console.error("✗ ERROR: Mammoth HTML does NOT contain merge fields!");
      console.log("Mammoth HTML sample (first 1000 chars):", mammothHtml.substring(0, 1000));
      console.log("This means the Word document might not have merge fields, or mammoth.js is not preserving them");
    }

    // Step 2: Use Puppeteer + docx-preview to get perfect styling
    console.log("Step 2: Using Puppeteer + docx-preview to get perfect styling...");
    const base64Word = wordBuffer.toString("base64");

    // Launch Puppeteer browser
    browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: chromium.defaultViewport,
      executablePath: await chromium.executablePath(),
      headless: true
    });

    const page = await browser.newPage();

    // Create HTML page that loads docx-preview and renders the Word document
    const htmlPage = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Word to HTML Conversion</title>
  <!-- Load JSZip (required dependency for docx-preview) -->
  <script src="https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js"></script>
  <!-- Load docx-preview (latest stable version) -->
  <script src="https://cdn.jsdelivr.net/npm/docx-preview@latest/dist/docx-preview.min.js"></script>
  <style>
    body {
      margin: 0;
      padding: 20px;
      background: white;
    }
    #container {
      width: 100%;
    }
    .docx-wrapper {
      background: white !important;
      padding: 0 !important;
    }
  </style>
</head>
<body>
  <div id="container"></div>
  <script>
    (async function() {
      try {
        // Wait for libraries to load
        if (typeof JSZip === 'undefined') {
          throw new Error('JSZip not loaded');
        }
        if (typeof docx === 'undefined' && typeof docxjs === 'undefined') {
          throw new Error('docx-preview not loaded');
        }

        // Get docx-preview library
        let docxLib = null;
        if (window.docx && typeof window.docx.renderAsync === 'function') {
          docxLib = window.docx;
        } else if (window.docxjs && typeof window.docxjs.renderAsync === 'function') {
          docxLib = window.docxjs;
        } else if (window.docxPreview && typeof window.docxPreview.renderAsync === 'function') {
          docxLib = window.docxPreview;
        } else {
          const possibleNames = ['docx', 'docxjs', 'docxPreview', 'docxPreviewjs'];
          for (const name of possibleNames) {
            if (window[name] && typeof window[name].renderAsync === 'function') {
              docxLib = window[name];
              break;
            }
          }
        }
        
        if (!docxLib || typeof docxLib.renderAsync !== 'function') {
          throw new Error('docx-preview renderAsync method not found');
        }

        // Decode base64 Word document to ArrayBuffer
        const base64Word = '${base64Word}';
        const binaryString = atob(base64Word);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        const arrayBuffer = bytes.buffer;

        // Render Word document using docx-preview (preserves all formatting, styles, colors, alignment, images)
        const container = document.getElementById('container');
        await docxLib.renderAsync(arrayBuffer, container, null, {
          useBase64URL: true,
          className: 'docx',
          inWrapper: true
        });

        // Wait for images to load
        const images = container.querySelectorAll('img');
        if (images.length > 0) {
          await new Promise(resolve => setTimeout(resolve, 500));
          await Promise.all(Array.from(images).map((img) => {
            return new Promise((resolve) => {
              if (img.complete && img.naturalHeight !== 0) {
                resolve();
              } else {
                img.onload = resolve;
                img.onerror = resolve;
                setTimeout(resolve, 3000);
              }
            });
          }));
        }

        // Extract styled HTML and styles from docx-preview
        const wrapper = container.querySelector('.docx-wrapper') || container;
        const styledHtml = wrapper.innerHTML;
        
        // Extract all styles (inline and from style tags)
        const allStyles = [];
        const styleTags = document.querySelectorAll('style');
        styleTags.forEach(style => {
          if (style.textContent) {
            allStyles.push(style.textContent);
          }
        });
        
        // Get computed styles for key elements to preserve colors, fonts, alignment
        const styleData = {
          html: styledHtml,
          styles: allStyles.join('\\n'),
          // Extract inline styles from key elements
          elementStyles: {}
        };
        
        // Extract styles from sample elements
        const sampleElements = wrapper.querySelectorAll('p, h1, h2, h3, table, td, th, span, div');
        sampleElements.forEach((el, index) => {
          if (index < 20) { // Limit to first 20 elements
            const computedStyle = window.getComputedStyle(el);
            styleData.elementStyles[el.tagName.toLowerCase() + '_' + index] = {
              color: computedStyle.color,
              backgroundColor: computedStyle.backgroundColor,
              fontSize: computedStyle.fontSize,
              fontWeight: computedStyle.fontWeight,
              textAlign: computedStyle.textAlign,
              fontFamily: computedStyle.fontFamily,
              margin: computedStyle.margin,
              padding: computedStyle.padding,
              border: computedStyle.border
            };
          }
        });

        // Signal completion with both styled HTML and style data
        window.conversionComplete = true;
        window.conversionError = null;
        window.styledHtml = styledHtml;
        window.styleData = styleData;
      } catch (error) {
        console.error('Conversion error:', error);
        window.conversionComplete = true;
        window.conversionError = error.message;
      }
    })();
  </script>
</body>
</html>`;

    // Load the HTML page
    await page.setContent(htmlPage, { waitUntil: "networkidle0" });

    // Wait for conversion to complete
    await page.waitForFunction(() => window.conversionComplete === true, { timeout: 60000 });

    // Check for errors
    const error = await page.evaluate(() => window.conversionError);
    if (error) {
      throw new Error("Conversion failed in browser: " + error);
    }

    // Extract styled HTML and style data from docx-preview
    const { styledHtml, styleData } = await page.evaluate(() => {
      return {
        styledHtml: window.styledHtml,
        styleData: window.styleData
      };
    });

    if (!styledHtml || styledHtml.trim().length === 0) {
      throw new Error("Rendered HTML is empty");
    }

    // Step 3: Merge mammoth HTML (with merge fields) into docx-preview styled structure
    console.log("Step 3: Merging mammoth HTML (with merge fields) into docx-preview styled structure...");
    
    // Use Puppeteer to merge the HTMLs intelligently
    // Strategy: Use docx-preview HTML as base (perfect styling), inject merge fields from mammoth
    const mergedHtml = await page.evaluate((mammothHtmlContent, styledHtmlContent, styleDataContent) => {
      // Create temporary containers
      const mammothDiv = document.createElement('div');
      mammothDiv.innerHTML = mammothHtmlContent;
      
      const styledDiv = document.createElement('div');
      styledDiv.innerHTML = styledHtmlContent;
      
      // Verify mammoth HTML has merge fields - CRITICAL CHECK
      const mammothMergeFields = mammothHtmlContent.match(/\{\{[^}]+\}\}/g);
      if (!mammothMergeFields || mammothMergeFields.length === 0) {
        console.error('ERROR: Mammoth HTML does not contain merge fields! This is a critical issue.');
        console.log('Mammoth HTML sample (first 500 chars):', mammothHtmlContent.substring(0, 500));
        // Return styled HTML anyway (better than nothing, but merge fields won't work)
        return {
          html: styledWrapper.innerHTML,
          styles: docxStyles,
          mergeFieldsFound: 0
        };
      } else {
        console.log(`✓ Mammoth HTML contains ${mammothMergeFields.length} merge fields:`, mammothMergeFields.slice(0, 5));
      }
      
      // Extract styles from docx-preview
      const docxStyles = styleDataContent.styles || '';
      
      // Get styled wrapper from docx-preview - THIS HAS PERFECT STYLING
      const styledWrapper = styledDiv.querySelector('.docx-wrapper') || styledDiv.querySelector('section.docx') || styledDiv;
      
      // Check if docx-preview HTML already has merge fields
      const styledHtmlText = styledWrapper.innerHTML;
      const hasMergeFieldsInStyled = styledHtmlText.match(/\{\{[^}]+\}\}/g);
      
      if (hasMergeFieldsInStyled && hasMergeFieldsInStyled.length > 0) {
        // docx-preview preserved merge fields! Use it directly - PERFECT STYLING PRESERVED
        console.log('docx-preview HTML already contains merge fields, using it directly with perfect styling');
        return {
          html: styledWrapper.innerHTML, // This has ALL inline styles from docx-preview
          styles: docxStyles,
          mergeFieldsFound: hasMergeFieldsInStyled.length
        };
      }
      
      // Also check the full styled HTML (not just wrapper) for merge fields
      const fullStyledHtml = styledDiv.innerHTML;
      const hasMergeFieldsInFull = fullStyledHtml.match(/\{\{[^}]+\}\}/g);
      if (hasMergeFieldsInFull && hasMergeFieldsInFull.length > 0) {
        console.log('docx-preview full HTML contains merge fields, using it directly');
        return {
          html: fullStyledHtml, // Use full HTML to preserve wrapper structure
          styles: docxStyles,
          mergeFieldsFound: hasMergeFieldsInFull.length
        };
      }
      
      // docx-preview didn't preserve merge fields
      // BETTER APPROACH: Use docx-preview HTML as BASE (perfect styling) and inject merge fields from mammoth
      console.log('docx-preview HTML missing merge fields. Using docx-preview HTML as base and injecting merge fields from mammoth...');
      console.log(`Mammoth HTML has ${mammothMergeFields.length} merge fields - these will be injected into styled HTML`);
      
      // CRITICAL: Use docx-preview HTML structure (perfect styling) as the base
      // Extract text content from mammoth HTML (has merge fields) and inject into docx-preview structure
      const article = styledWrapper.querySelector('article') || styledWrapper.querySelector('section') || styledWrapper;
      
      // Create container for mammoth HTML to extract text
      const mammothContainer = document.createElement('div');
      mammothContainer.innerHTML = mammothHtmlContent;
      
      // Extract all text nodes from mammoth HTML (these contain merge fields)
      const mammothTextNodes = [];
      const mammothWalker = document.createTreeWalker(
        mammothContainer,
        NodeFilter.SHOW_TEXT,
        null,
        false
      );
      let mammothNode;
      while (mammothNode = mammothWalker.nextNode()) {
        const text = mammothNode.textContent;
        if (text && text.trim()) {
          mammothTextNodes.push(text);
        }
      }
      
      console.log(`Extracted ${mammothTextNodes.length} text nodes from mammoth HTML (with merge fields)`);
      
      // Extract all text nodes from docx-preview HTML (perfect styling, but no merge fields)
      const styledTextNodes = [];
      const styledTextNodesMap = new Map(); // Map to store node -> parent element for style preservation
      const styledWalker = document.createTreeWalker(
        article,
        NodeFilter.SHOW_TEXT,
        null,
        false
      );
      let styledNode;
      while (styledNode = styledWalker.nextNode()) {
        const text = styledNode.textContent;
        if (text && text.trim()) {
          styledTextNodes.push({
            node: styledNode,
            parent: styledNode.parentElement,
            originalText: text
          });
        }
      }
      
      console.log(`Found ${styledTextNodes.length} text nodes in docx-preview HTML (with perfect styling)`);
      
      // Replace text content in docx-preview HTML with text from mammoth HTML
      // This preserves ALL inline styles, colors, fonts from docx-preview
      let mammothIndex = 0;
      styledTextNodes.forEach(({ node, parent }) => {
        if (mammothIndex < mammothTextNodes.length && parent) {
          // Replace text content - parent element's inline styles are preserved automatically
          node.textContent = mammothTextNodes[mammothIndex];
          mammothIndex++;
        }
      });
      
      // If there are remaining mammoth text nodes, append them with similar styling
      if (mammothIndex < mammothTextNodes.length) {
        console.log(`Appending ${mammothTextNodes.length - mammothIndex} additional text nodes from mammoth...`);
        
        // Get styles from last paragraph to maintain consistency
        const lastP = Array.from(article.querySelectorAll('p')).pop();
        const baseStyle = lastP ? lastP.style.cssText : '';
        const baseClass = lastP ? lastP.className : '';
        
        for (let i = mammothIndex; i < mammothTextNodes.length; i++) {
          const p = document.createElement('p');
          if (baseStyle) p.style.cssText = baseStyle;
          if (baseClass) p.className = baseClass;
          p.textContent = mammothTextNodes[i];
          article.appendChild(p);
        }
      }
      
      // The article now has docx-preview's perfect styling with mammoth's merge fields
      // No need to replace content - we've already updated the text nodes
      
      // Get merged HTML - should have merge fields from mammoth AND perfect styling from docx-preview
      // The styledWrapper contains the complete docx-preview HTML structure with all inline styles
      // We've replaced text nodes, so merge fields are injected while preserving all styling
      let mergedHtml = styledWrapper.innerHTML;
      
      // CRITICAL: Verify that inline styles are still present in the HTML
      // Count elements with inline styles to ensure they're preserved
      const elementsWithStyles = styledWrapper.querySelectorAll('[style]');
      console.log(`✓ Preserved ${elementsWithStyles.length} elements with inline styles from docx-preview`);
      
      // CRITICAL: Fix split merge fields - docx-preview may split {{FieldName}} across <span> elements
      // We need to reconstruct them into single text nodes so Apex can replace them
      console.log('Fixing split merge fields (reconstructing merge fields split across HTML elements)...');
      
      // Function to reconstruct split merge fields
      const fixSplitMergeFields = (container) => {
        // Get all text nodes
        const textNodes = [];
        const walker = document.createTreeWalker(
          container,
          NodeFilter.SHOW_TEXT,
          null,
          false
        );
        let node;
        while (node = walker.nextNode()) {
          textNodes.push(node);
        }
        
        // Look for merge fields split across multiple text nodes
        for (let i = 0; i < textNodes.length; i++) {
          const textNode = textNodes[i];
          const text = textNode.textContent;
          
          // Check if this node starts a merge field
          if (text.includes('{{') || text.includes('[[')) {
            // Collect adjacent text nodes that might contain the rest of the merge field
            let mergeFieldParts = [];
            let currentIndex = i;
            let foundEnd = false;
            
            // Collect text from current and following nodes (up to 5 nodes)
            while (currentIndex < textNodes.length && currentIndex - i <= 5 && !foundEnd) {
              const currentNode = textNodes[currentIndex];
              const currentText = currentNode.textContent;
              
              mergeFieldParts.push({ node: currentNode, text: currentText, index: currentIndex });
              
              // Check if we found the end
              if (currentText.includes('}}') || currentText.includes(']]')) {
                foundEnd = true;
                break;
              }
              
              currentIndex++;
            }
            
            // If we found a split merge field (across multiple nodes)
            if (foundEnd && mergeFieldParts.length > 1) {
              const combinedText = mergeFieldParts.map(p => p.text).join('');
              
              // Check if combined text forms a complete merge field
              const mergeFieldMatch = combinedText.match(/(\{\{|\]\])(.+?)(\}\}|\]\])/);
              
              if (mergeFieldMatch) {
                const fullMatch = mergeFieldMatch[0];
                const fieldName = mergeFieldMatch[2].trim();
                
                console.log(`Found split merge field: "${fieldName}" - reconstructing...`);
                
                // Reconstruct: put the complete merge field in the first node, clear others
                const firstPart = mergeFieldParts[0];
                const beforeMatch = combinedText.substring(0, combinedText.indexOf(fullMatch));
                const afterMatch = combinedText.substring(combinedText.indexOf(fullMatch) + fullMatch.length);
                
                // Update first node with complete merge field
                firstPart.node.textContent = beforeMatch + fullMatch + afterMatch;
                
                // Clear other nodes (they're now part of the first node)
                for (let j = 1; j < mergeFieldParts.length; j++) {
                  mergeFieldParts[j].node.textContent = '';
                }
                
                // Skip processed nodes
                i = currentIndex;
              }
            }
          }
        }
      };
      
      // Fix split merge fields in the merged HTML
      fixSplitMergeFields(styledWrapper);
      
      // Get updated HTML after fixing split merge fields
      mergedHtml = styledWrapper.innerHTML;
      
      // CRITICAL: Verify merge fields are present and properly formatted
      const mergeFieldCheck = mergedHtml.match(/\{\{[^}]+\}\}/g);
      console.log(`After fixing split merge fields: Found ${mergeFieldCheck ? mergeFieldCheck.length : 0} merge fields`);
      
      // Also check for split merge fields (they should be fixed now)
      const splitMergeFieldCheck = mergedHtml.match(/\{\{<\/span><span>[^}]+\}\}/g);
      if (splitMergeFieldCheck && splitMergeFieldCheck.length > 0) {
        console.warn(`WARNING: Still found ${splitMergeFieldCheck.length} split merge fields that couldn't be fixed`);
      }
      
      if (!mergeFieldCheck || mergeFieldCheck.length === 0) {
        console.error('ERROR: Merge fields lost! Using mammoth HTML directly (no style application).');
        // This should never happen, but if it does, return mammoth HTML as-is
        mergedHtml = mammothHtmlContent;
        const directCheck = mergedHtml.match(/\{\{[^}]+\}\}/g);
        if (directCheck && directCheck.length > 0) {
          console.log(`✓ Direct mammoth HTML has ${directCheck.length} merge fields`);
        } else {
          console.error('CRITICAL: Even direct mammoth HTML has no merge fields!');
        }
      } else {
        console.log(`✓ SUCCESS: ${mergeFieldCheck.length} merge fields preserved with docx-preview styling`);
        console.log('Sample fixed merge fields:', mergeFieldCheck.slice(0, 5));
      }
      
      // Final check before returning - also fix split merge fields at string level
      // CRITICAL: Fix split merge fields like {{</span><span>Field</span><span>}} -> {{Field}}
      // This must happen AFTER we get the HTML from Puppeteer but BEFORE returning
      let fixedHtml = mergedHtml;
      let fixedCount = 0;
      
      console.log('Starting string-level fix for split merge fields...');
      
      // AGGRESSIVE FIX: Match the exact pattern we see in logs: {{</span><span>FieldName</span><span>}}
      // Use a more flexible pattern that captures everything between {{ and }} across spans
      const splitPattern1 = /\{\{<\/span><span>([^<]+)<\/span><span>\}\}/g;
      fixedHtml = fixedHtml.replace(splitPattern1, (match, fieldName) => {
        fixedCount++;
        const fixed = `{{${fieldName}}}`;
        console.log(`Fixed pattern 1: "${match}" -> "${fixed}"`);
        return fixed;
      });
      
      // Pattern 2: {{</span><span>Field</span><span>Part</span><span>}} (multiple spans)
      const splitPattern2 = /\{\{<\/span>(<span>([^<]+)<\/span>)+<span>\}\}/g;
      fixedHtml = fixedHtml.replace(splitPattern2, (match) => {
        // Extract all field parts
        const spanMatches = match.match(/<span>([^<]+)<\/span>/g);
        if (spanMatches && spanMatches.length > 0) {
          const fieldParts = spanMatches.map(span => span.replace(/<\/?span>/g, ''));
          const fieldName = fieldParts.join('');
          const fixed = `{{${fieldName}}}`;
          fixedCount++;
          console.log(`Fixed pattern 2: "${match.substring(0, 60)}..." -> "${fixed}"`);
          return fixed;
        }
        return match;
      });
      
      // Pattern 3: {{<span>Field</span><span>Name</span>}} (no closing span before }})
      const splitPattern3 = /\{\{<span>([^<]+)<\/span>(<span>([^<]+)<\/span>)*\}\}/g;
      fixedHtml = fixedHtml.replace(splitPattern3, (match) => {
        const spanMatches = match.match(/<span>([^<]+)<\/span>/g);
        if (spanMatches && spanMatches.length > 0) {
          const fieldParts = spanMatches.map(span => span.replace(/<\/?span>/g, ''));
          const fieldName = fieldParts.join('');
          const fixed = `{{${fieldName}}}`;
          fixedCount++;
          console.log(`Fixed pattern 3: "${match.substring(0, 60)}..." -> "${fixed}"`);
          return fixed;
        }
        return match;
      });
      
      // Pattern 4: More complex - handle cases like {{</span><span>OpportunityLineItems</span><span>[0]</span><span>.Product2.Name</span><span>}}
      // This uses a greedy approach to match everything between {{ and }} that contains spans
      const splitPattern4 = /\{\{<\/span>((?:<span>[^<]+<\/span>)+)<span>\}\}/g;
      fixedHtml = fixedHtml.replace(splitPattern4, (match, spansContent) => {
        const spanMatches = spansContent.match(/<span>([^<]+)<\/span>/g);
        if (spanMatches && spanMatches.length > 0) {
          const fieldParts = spanMatches.map(span => span.replace(/<\/?span>/g, ''));
          const fieldName = fieldParts.join('');
          const fixed = `{{${fieldName}}}`;
          fixedCount++;
          console.log(`Fixed pattern 4: "${match.substring(0, 80)}..." -> "${fixed}"`);
          return fixed;
        }
        return match;
      });
      
      if (fixedCount > 0) {
        console.log(`✓ Fixed ${fixedCount} split merge fields at string level`);
        mergedHtml = fixedHtml;
        
        // Verify fix worked
        const afterFixCheck = mergedHtml.match(/\{\{[^}]+\}\}/g);
        const splitAfterFix = mergedHtml.match(/\{\{<\/span><span>/g);
        if (splitAfterFix && splitAfterFix.length > 0) {
          console.warn(`WARNING: Still found ${splitAfterFix.length} split merge fields after string-level fix`);
          console.log('Sample remaining split fields:', splitAfterFix.slice(0, 3));
        } else {
          console.log(`✓ All split merge fields fixed - found ${afterFixCheck ? afterFixCheck.length : 0} proper merge fields`);
          if (afterFixCheck && afterFixCheck.length > 0) {
            console.log('Sample fixed merge fields:', afterFixCheck.slice(0, 5));
          }
        }
      } else {
        // Check if there are split merge fields that weren't caught
        const splitCheck = mergedHtml.match(/\{\{<\/span><span>/g);
        if (splitCheck && splitCheck.length > 0) {
          console.warn(`WARNING: Found ${splitCheck.length} split merge fields but fix didn't catch them!`);
          console.log('Sample split pattern:', splitCheck[0]);
          // Try one more aggressive fix - remove all span tags inside merge fields
          fixedHtml = mergedHtml.replace(/\{\{([^}]*<span>[^<]+<\/span>[^}]*)\}\}/g, (match, content) => {
            const cleaned = content.replace(/<\/?span>/g, '');
            return `{{${cleaned}}}`;
          });
          if (fixedHtml !== mergedHtml) {
            mergedHtml = fixedHtml;
            console.log('Applied aggressive cleanup - removed all span tags inside merge fields');
          }
        }
      }
      
      const finalCheck = mergedHtml.match(/\{\{[^}]+\}\}/g) || [];
      const mergeFieldsCount = finalCheck.length;
      
      console.log(`Final merge result: ${mergeFieldsCount} merge fields found in merged HTML`);
      if (mergeFieldsCount > 0) {
        console.log('✓ Sample merge fields:', finalCheck.slice(0, 5));
      } else {
        console.warn('WARNING: No merge fields found in final merged HTML!');
        // Last resort: return mammoth HTML with docx-preview styles
        console.log('Using mammoth HTML as fallback with docx-preview wrapper...');
        const fallbackWrapper = styledWrapper.cloneNode(false);
        fallbackWrapper.innerHTML = mammothHtmlContent;
        mergedHtml = fallbackWrapper.innerHTML;
        const fallbackCheck = mergedHtml.match(/\{\{[^}]+\}\}/g) || [];
        if (fallbackCheck.length > 0) {
          console.log(`✓ Fallback successful: ${fallbackCheck.length} merge fields found`);
          return {
            html: mergedHtml,
            styles: docxStyles,
            mergeFieldsFound: fallbackCheck.length
          };
        } else {
          console.error('ERROR: Even fallback has no merge fields!');
          // Return mammoth HTML anyway
          return {
            html: mammothHtmlContent,
            styles: docxStyles,
            mergeFieldsFound: 0
          };
        }
      }
      
      // Get the final HTML with ALL inline styles preserved
      // This is the docx-preview HTML structure with merge fields injected
      const finalMergedHtml = styledWrapper.innerHTML;
      
      return {
        html: finalMergedHtml,
        styles: docxStyles,
        mergeFieldsFound: mergeFieldsCount
      };
    }, mammothHtml, styledHtml, styleData);

    // Extract final HTML one more time to ensure we have the latest version with all styles
    const finalHtmlWithStyles = await page.evaluate(() => {
      const wrapper = document.querySelector('.docx-wrapper') || document.querySelector('section.docx') || document.body;
      return wrapper.innerHTML;
    });

    // Update mergedHtml.html with the latest version (in case text node replacement updated it)
    if (finalHtmlWithStyles && finalHtmlWithStyles !== mergedHtml.html) {
      console.log('Updating merged HTML with latest styled version from browser...');
      mergedHtml.html = finalHtmlWithStyles;
    }

    // Close browser
    await browser.close();
    browser = null;

    // CRITICAL SERVER-SIDE FIX: Fix split merge fields in the HTML we got from Puppeteer
    // This runs on the Node.js server, not in the browser, so we can do aggressive string replacement
    console.log("Applying server-side fix for split merge fields...");
    let serverFixedHtml = mergedHtml.html;
    let serverFixedCount = 0;
    
    // Pattern 1: {{</span><span>FieldName</span><span>}}
    const serverPattern1 = /\{\{<\/span><span>([^<]+)<\/span><span>\}\}/g;
    serverFixedHtml = serverFixedHtml.replace(serverPattern1, (match, fieldName) => {
      serverFixedCount++;
      return `{{${fieldName}}}`;
    });
    
    // Pattern 2: Multiple spans {{</span><span>Field</span><span>Part</span><span>}}
    const serverPattern2 = /\{\{<\/span>((?:<span>[^<]+<\/span>)+)<span>\}\}/g;
    serverFixedHtml = serverFixedHtml.replace(serverPattern2, (match, spansContent) => {
      const spanMatches = spansContent.match(/<span>([^<]+)<\/span>/g);
      if (spanMatches && spanMatches.length > 0) {
        const fieldParts = spanMatches.map(span => span.replace(/<\/?span>/g, ''));
        const fieldName = fieldParts.join('');
        serverFixedCount++;
        return `{{${fieldName}}}`;
      }
      return match;
    });
    
    // Pattern 3: {{<span>Field</span><span>Name</span>}} (no closing span)
    const serverPattern3 = /\{\{<span>([^<]+)<\/span>(<span>([^<]+)<\/span>)*\}\}/g;
    serverFixedHtml = serverFixedHtml.replace(serverPattern3, (match) => {
      const spanMatches = match.match(/<span>([^<]+)<\/span>/g);
      if (spanMatches && spanMatches.length > 0) {
        const fieldParts = spanMatches.map(span => span.replace(/<\/?span>/g, ''));
        const fieldName = fieldParts.join('');
        serverFixedCount++;
        return `{{${fieldName}}}`;
      }
      return match;
    });
    
    // Pattern 4: Aggressive - remove ALL span tags inside any {{...}} that contains spans
    const serverPattern4 = /\{\{([^}]*<span>[^<]+<\/span>[^}]*)\}\}/g;
    serverFixedHtml = serverFixedHtml.replace(serverPattern4, (match, content) => {
      const cleaned = content.replace(/<\/?span>/g, '');
      serverFixedCount++;
      return `{{${cleaned}}}`;
    });
    
    if (serverFixedCount > 0) {
      console.log(`✓ Server-side fix: Fixed ${serverFixedCount} split merge fields`);
      mergedHtml.html = serverFixedHtml;
    }
    
    // Verify merge fields are still present and properly formatted
    const finalMergeFields = mergedHtml.html.match(/\{\{[^}]+\}\}/g) || [];
    const splitMergeFields = mergedHtml.html.match(/\{\{<\/span><span>/g) || [];
    
    console.log("Merge fields in merged HTML:", finalMergeFields.length, "fields");
    if (splitMergeFields.length > 0) {
      console.warn(`WARNING: Still found ${splitMergeFields.length} split merge fields after server-side fix!`);
      console.log("Sample split pattern:", splitMergeFields[0]);
    } else {
      console.log("✓ All merge fields are properly formatted (no split patterns found)");
      if (finalMergeFields.length > 0) {
        console.log("Sample merge fields:", finalMergeFields.slice(0, 5));
      }
    }
    
    // CRITICAL: If merge fields are missing, use mammoth HTML directly with docx-preview wrapper
    if (finalMergeFields.length === 0 && mergeFields.length > 0) {
      console.warn("CRITICAL: Merge fields were lost during merge. Using mammoth HTML directly with docx-preview wrapper.");
      console.log(`Original mammoth HTML had ${mergeFields.length} merge fields:`, mergeFields.slice(0, 5));
      
      // Use mammoth HTML directly (guarantees merge fields) with docx-preview wrapper structure
      const finalHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    ${mergedHtml.styles || ''}
    /* Additional styles to enhance mammoth HTML */
    body {
      font-family: 'Calibri', 'Arial', 'Helvetica', sans-serif;
      margin: 20px;
      color: #000000;
      background: white;
    }
    .docx-wrapper {
      background: white !important;
      padding: 0 !important;
    }
    * {
      box-sizing: border-box;
    }
    /* Preserve all inline styles from mammoth */
    *[style] {
      /* Inline styles take precedence */
    }
  </style>
</head>
<body>
  <div class="docx-wrapper">
    ${mammothHtml}
  </div>
</body>
</html>`;
      
      // Verify mammoth HTML still has merge fields
      const mammothCheck = finalHtml.match(/\{\{[^}]+\}\}/g) || [];
      console.log(`Fallback HTML contains ${mammothCheck.length} merge fields`);
      
      // Always use fallback if merge fields were lost
      console.log("Using fallback HTML with mammoth content (merge fields guaranteed)");
      if (format === "json") {
        const response = { success: true, html: finalHtml };
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        return res.json(response);
      } else {
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        return res.send(finalHtml);
      }
    }

    // Build final HTML with merged content and styles
    // CRITICAL: Preserve ALL inline styles from docx-preview - they're already in mergedHtml.html
    // Don't wrap in additional divs that might break styling
    let finalHtml;
    
    // Check if mergedHtml.html already has full HTML structure
    if (mergedHtml.html.trim().toLowerCase().startsWith('<!doctype') || 
        mergedHtml.html.trim().toLowerCase().startsWith('<html')) {
      // Already has full structure, just ensure styles are included
      finalHtml = mergedHtml.html;
      // Inject styles into head if not present
      if (!finalHtml.includes('<style>') && mergedHtml.styles) {
        finalHtml = finalHtml.replace('</head>', `<style>${mergedHtml.styles}</style></head>`);
      }
    } else {
      // Wrap in HTML structure, preserving all inline styles from mergedHtml.html
      finalHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    ${mergedHtml.styles || ''}
    /* Preserve all inline styles from docx-preview - they're in the HTML elements */
    body {
      font-family: 'Calibri', 'Arial', 'Helvetica', sans-serif;
      margin: 0;
      padding: 20px;
      color: #000000;
      background: white;
    }
    .docx-wrapper {
      background: white !important;
      padding: 0 !important;
    }
    /* CRITICAL: Don't override inline styles - they take precedence */
    * {
      box-sizing: border-box;
    }
    /* Ensure all colors, fonts, alignment from inline styles are preserved */
    *[style] {
      /* Inline styles will override any CSS rules */
    }
  </style>
</head>
<body>
  ${mergedHtml.html}
</body>
</html>`;
    }

    // FINAL VERIFICATION: Ensure merge fields are present in final HTML
    const finalCheck = finalHtml.match(/\{\{[^}]+\}\}/g) || [];
    console.log("Word to HTML conversion successful. HTML length:", finalHtml.length, "characters");
    console.log("Final merge fields count:", finalCheck.length, "fields");
    
    if (finalCheck.length > 0) {
      console.log("✓ SUCCESS: Merge fields preserved:", finalCheck.slice(0, 5));
      console.log("Styling from docx-preview: Applied");
    } else if (mergeFields.length > 0) {
      console.error("✗ ERROR: Merge fields were lost! Original had", mergeFields.length, "fields");
      console.log("This should not happen - fallback should have been used");
    }

    if (format === "json") {
      const response = {
        success: true,
        html: finalHtml
      };
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      return res.json(response);
    } else {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      return res.send(finalHtml);
    }
  } catch (error) {
    console.error("Word to HTML conversion error:", error);
    
    if (browser) {
      try {
        await browser.close();
      } catch (closeError) {
        console.error("Error closing browser:", closeError);
      }
    }
    
    return res.status(500).json({ 
      success: false, 
      error: "Failed to convert Word to HTML", 
      details: error.message 
    });
  }
  });

app.post("/convert", async (req, res) => {
  try {
    const { html } = req.body;
    const format = req.query.format?.toLowerCase() === "json" ? "json" : "pdf";

    if (!html) {
      return res.status(400).json({ success: false, error: "HTML content is required" });
    }

    const browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: chromium.defaultViewport,
      executablePath: await chromium.executablePath(),
      headless: true
    });

    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: ["load", "domcontentloaded"] });

    const pdfBuffer = await page.pdf({ format: "A4", printBackground: true });
    await browser.close();

    if (format === "json") {
      // Validate PDF buffer first
      console.log("PDF buffer size:", pdfBuffer.length);
      console.log("PDF buffer type:", typeof pdfBuffer);
      console.log("PDF buffer is Buffer:", Buffer.isBuffer(pdfBuffer));
      
      // CRITICAL FIX: Puppeteer returns a Uint8Array, not a Node.js Buffer
      // Convert to proper Node.js Buffer first, then to base64
      let nodeBuffer;
      if (Buffer.isBuffer(pdfBuffer)) {
        nodeBuffer = pdfBuffer;
      } else if (pdfBuffer instanceof Uint8Array) {
        nodeBuffer = Buffer.from(pdfBuffer);
      } else {
        // Fallback: try to convert whatever it is
        nodeBuffer = Buffer.from(pdfBuffer);
      }
      
      console.log("Node buffer size:", nodeBuffer.length);
      console.log("Node buffer is Buffer:", Buffer.isBuffer(nodeBuffer));
      
      // Now convert to base64 properly
      let base64Pdf = nodeBuffer.toString("base64");
      
      // Debug: Check base64 BEFORE any cleaning
      console.log("Base64 BEFORE cleaning - length:", base64Pdf.length);
      console.log("Base64 BEFORE cleaning - first 100 chars:", base64Pdf.substring(0, 100));
      console.log("Base64 BEFORE cleaning - contains letters:", /[A-Za-z]/.test(base64Pdf));
      console.log("Base64 BEFORE cleaning - contains +:", base64Pdf.includes('+'));
      console.log("Base64 BEFORE cleaning - contains /:", base64Pdf.includes('/'));
      console.log("Base64 BEFORE cleaning - contains =:", base64Pdf.includes('='));
      
      // Remove any invalid characters: whitespace, commas, line breaks
      // BUT be careful - only remove actual invalid chars, not valid base64 chars
      base64Pdf = base64Pdf.replace(/[\s\n\r\t]/g, ""); // Only remove whitespace
      
      // Debug: Check base64 AFTER cleaning
      console.log("Base64 AFTER cleaning - length:", base64Pdf.length);
      console.log("Base64 AFTER cleaning - first 100 chars:", base64Pdf.substring(0, 100));
      console.log("Base64 AFTER cleaning - contains letters:", /[A-Za-z]/.test(base64Pdf));
      
      // CRITICAL VALIDATION: Base64 MUST contain letters (A-Z or a-z)
      // If it doesn't, the PDF buffer or base64 conversion is corrupted
      if (!/[A-Za-z]/.test(base64Pdf)) {
        console.error("ERROR: Base64 string contains NO letters! This is invalid.");
        console.error("First 200 chars:", base64Pdf.substring(0, 200));
        console.error("All chars are digits:", /^\d+$/.test(base64Pdf));
        return res.status(500).json({ 
          success: false, 
          error: "Generated base64 is invalid - contains no letters. PDF buffer may be corrupted." 
        });
      }
      
      // Ensure base64 is sent as a proper string in JSON
      const response = {
        success: true,
        pdf: String(base64Pdf) // Explicitly convert to string
      };
      
      // Debug: Verify response object
      console.log("Response pdf field type:", typeof response.pdf);
      console.log("Response pdf field length:", response.pdf.length);
      console.log("Response pdf field (first 100 chars):", response.pdf.substring(0, 100));
      
      // Set content type and send with explicit JSON stringification
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      const jsonString = JSON.stringify(response);
      console.log("Final JSON (first 200 chars):", jsonString.substring(0, 200));
      
      return res.send(jsonString);
    } else {
      res.set({ "Content-Type": "application/pdf", "Content-Length": pdfBuffer.length });
      return res.send(pdfBuffer);
    }
  } catch (error) {
    console.error("PDF error:", error);
    return res.status(500).json({ success: false, error: "Failed to generate PDF", details: error.message });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on ${PORT}`));
