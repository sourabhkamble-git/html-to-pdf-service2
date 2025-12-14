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
      // SIMPLIFIED APPROACH: Always use mammoth HTML (guarantees merge fields) and apply docx-preview styles
      console.log('docx-preview HTML missing merge fields. Using mammoth HTML with docx-preview styles applied...');
      console.log(`Mammoth HTML has ${mammothMergeFields.length} merge fields - these will be preserved`);
      
      // Use mammoth HTML as base (GUARANTEES merge fields) and apply styles from docx-preview
      const article = styledWrapper.querySelector('article') || styledWrapper.querySelector('section') || styledWrapper;
      
      // Create container for mammoth HTML
      const mammothContainer = document.createElement('div');
      mammothContainer.innerHTML = mammothHtmlContent;
      
      // Get styled elements from docx-preview
      const styledElements = Array.from(article.querySelectorAll('*'));
      const mammothElements = Array.from(mammothContainer.querySelectorAll('*'));
      
      console.log(`Applying styles from ${styledElements.length} styled elements to ${mammothElements.length} mammoth elements`);
      
      // Create style map: tag -> array of styles (by position)
      const styleMap = new Map();
      styledElements.forEach((el) => {
        const tag = el.tagName;
        if (!styleMap.has(tag)) {
          styleMap.set(tag, []);
        }
        styleMap.get(tag).push({
          style: el.style.cssText,
          className: el.className,
          attributes: Array.from(el.attributes).filter(attr => 
            ['style', 'class', 'align', 'valign', 'width', 'height', 'colspan', 'rowspan', 'bgcolor'].includes(attr.name)
          ).reduce((acc, attr) => {
            acc[attr.name] = attr.value;
            return acc;
          }, {})
        });
      });
      
      // Apply styles to mammoth elements by matching tag and position
      const tagIndices = new Map();
      mammothElements.forEach((mammothEl) => {
        const tag = mammothEl.tagName;
        if (!tagIndices.has(tag)) {
          tagIndices.set(tag, 0);
        }
        const index = tagIndices.get(tag);
        tagIndices.set(tag, index + 1);
        
        // Get style for this tag at this position
        if (styleMap.has(tag) && styleMap.get(tag).length > index) {
          const styleData = styleMap.get(tag)[index];
          
          // Apply inline style
          if (styleData.style) {
            mammothEl.style.cssText = styleData.style;
          }
          
          // Apply class
          if (styleData.className) {
            mammothEl.className = styleData.className;
          }
          
          // Apply attributes
          Object.entries(styleData.attributes).forEach(([name, value]) => {
            mammothEl.setAttribute(name, value);
          });
        } else if (styleMap.has(tag) && styleMap.get(tag).length > 0) {
          // Use first style of this tag type as fallback
          const styleData = styleMap.get(tag)[0];
          if (styleData.style) {
            mammothEl.style.cssText = styleData.style;
          }
          if (styleData.className) {
            mammothEl.className = styleData.className;
          }
        }
      });
      
      // Replace article/wrapper content with styled mammoth HTML
      if (article !== styledWrapper) {
        const newArticle = document.createElement(article.tagName || 'article');
        if (article.className) newArticle.className = article.className;
        if (article.style.cssText) newArticle.style.cssText = article.style.cssText;
        newArticle.innerHTML = mammothContainer.innerHTML;
        article.innerHTML = newArticle.innerHTML;
      } else {
        // Preserve wrapper styles
        const wrapperClone = styledWrapper.cloneNode(false);
        wrapperClone.innerHTML = mammothContainer.innerHTML;
        styledWrapper.innerHTML = wrapperClone.innerHTML;
      }
      
      // Get merged HTML - should have merge fields from mammoth
      let mergedHtml = styledWrapper.innerHTML;
      
      // CRITICAL: Verify merge fields are present (they should be, we used mammoth HTML)
      const mergeFieldCheck = mergedHtml.match(/\{\{[^}]+\}\}/g);
      console.log(`After style application: Found ${mergeFieldCheck ? mergeFieldCheck.length : 0} merge fields`);
      
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
      }
      
      // Final check before returning
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
      
      return {
        html: mergedHtml,
        styles: docxStyles,
        mergeFieldsFound: mergeFieldsCount
      };
    }, mammothHtml, styledHtml, styleData);

    // Close browser
    await browser.close();
    browser = null;

    // Verify merge fields are still present
    const finalMergeFields = mergedHtml.html.match(/\{\{[^}]+\}\}/g) || [];
    console.log("Merge fields in merged HTML:", finalMergeFields.length, "fields");
    
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
