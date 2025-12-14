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
      console.log("Sample merge fields:", mergeFields.slice(0, 5));
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
    // Strategy: Inject mammoth HTML (with merge fields) into docx-preview's styled structure
    const mergedHtml = await page.evaluate((mammothHtmlContent, styledHtmlContent, styleDataContent) => {
      // Create temporary containers
      const mammothDiv = document.createElement('div');
      mammothDiv.innerHTML = mammothHtmlContent;
      
      const styledDiv = document.createElement('div');
      styledDiv.innerHTML = styledHtmlContent;
      
      // Extract styles from docx-preview
      const docxStyles = styleDataContent.styles || '';
      
      // Get styled wrapper from docx-preview
      const styledWrapper = styledDiv.querySelector('.docx-wrapper') || styledDiv.querySelector('section.docx') || styledDiv;
      
      // Find the article or main content container in styled HTML
      let contentContainer = styledWrapper.querySelector('article');
      if (!contentContainer) {
        // Create article if it doesn't exist
        contentContainer = document.createElement('article');
        styledWrapper.appendChild(contentContainer);
      }
      
      // Inject mammoth HTML (with merge fields) into the styled structure
      // This preserves docx-preview's wrapper and styling while using mammoth's content
      contentContainer.innerHTML = mammothHtmlContent;
      
      // Now we need to apply docx-preview's inline styles to the mammoth content
      // Extract inline styles from original styled elements and apply to corresponding mammoth elements
      const originalStyledElements = styledDiv.querySelectorAll('p, h1, h2, h3, table, td, th, span, div, li');
      const mammothElements = contentContainer.querySelectorAll('p, h1, h2, h3, table, td, th, span, div, li');
      
      // Apply styles from original to mammoth elements (by position/index)
      originalStyledElements.forEach((originalEl, index) => {
        if (index < mammothElements.length) {
          const mammothEl = mammothElements[index];
          // Copy inline styles from original to mammoth
          if (originalEl.style.cssText) {
            mammothEl.style.cssText = originalEl.style.cssText;
          }
          // Copy class names
          if (originalEl.className) {
            mammothEl.className = originalEl.className;
          }
        }
      });
      
      // Get merged HTML
      let mergedHtml = styledWrapper.innerHTML;
      
      // Verify merge fields are still present
      const mergeFieldCheck = mergedHtml.match(/\{\{[^}]+\}\}/g);
      
      return {
        html: mergedHtml,
        styles: docxStyles,
        mergeFieldsFound: mergeFieldCheck ? mergeFieldCheck.length : 0
      };
    }, mammothHtml, styledHtml, styleData);

    // Close browser
    await browser.close();
    browser = null;

    // Verify merge fields are still present
    const finalMergeFields = mergedHtml.html.match(/\{\{[^}]+\}\}/g) || [];
    console.log("Merge fields in merged HTML:", finalMergeFields.length, "fields");
    
    if (finalMergeFields.length === 0 && mergeFields.length > 0) {
      console.warn("Warning: Merge fields were lost during merge. Using mammoth HTML with enhanced styling.");
      // Fallback: Use mammoth HTML with docx-preview styles applied
      const finalHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    ${mergedHtml.styles}
    /* Additional styles to enhance mammoth HTML */
    body {
      font-family: 'Calibri', 'Arial', 'Helvetica', sans-serif;
      margin: 20px;
      color: #000000;
    }
    .docx-wrapper {
      background: white !important;
      padding: 0 !important;
    }
    * {
      box-sizing: border-box;
    }
  </style>
</head>
<body>
  <div class="docx-wrapper">
    ${mammothHtml}
  </div>
</body>
</html>`;
      
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
    const finalHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    ${mergedHtml.styles}
    /* Ensure merge fields are visible */
    body {
      font-family: 'Calibri', 'Arial', 'Helvetica', sans-serif;
      margin: 20px;
      color: #000000;
    }
    .docx-wrapper {
      background: white !important;
      padding: 0 !important;
    }
    * {
      box-sizing: border-box;
    }
  </style>
</head>
<body>
  <div class="docx-wrapper">
    ${mergedHtml.html}
  </div>
</body>
</html>`;

    console.log("Word to HTML conversion successful using hybrid approach. HTML length:", finalHtml.length, "characters");
    console.log("Merge fields preserved:", finalMergeFields.length, "fields");
    console.log("Styling from docx-preview: Applied");

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
