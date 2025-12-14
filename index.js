import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import chromium from "@sparticuz/chromium";
import puppeteer from "puppeteer-core";

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: "50mb" }));

app.get("/", (req, res) => {
    res.send("HTML to PDF service is running. Use POST /convert with HTML content, or POST /convert-word-to-html with Word document.");
  });

// New endpoint for Word to HTML conversion using Puppeteer + docx-preview
// This approach preserves ALL Word formatting, styles, colors, alignment, and images
// exactly as they appear in the original document (same as LWC component)
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

    console.log("Converting Word to HTML using Puppeteer + docx-preview. File size:", wordBuffer.length, "bytes");

    // Convert buffer to base64 data URL for embedding in HTML
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
    /* Remove docx-preview's default grey background */
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

        // Get docx-preview library (it might be exposed as 'docx', 'docxjs', or 'docxPreview')
        // docx-preview typically exposes itself as window.docx or window.docxjs
        let docxLib = null;
        if (window.docx && typeof window.docx.renderAsync === 'function') {
          docxLib = window.docx;
        } else if (window.docxjs && typeof window.docxjs.renderAsync === 'function') {
          docxLib = window.docxjs;
        } else if (window.docxPreview && typeof window.docxPreview.renderAsync === 'function') {
          docxLib = window.docxPreview;
        } else {
          // Try to find it in any global variable
          const possibleNames = ['docx', 'docxjs', 'docxPreview', 'docxPreviewjs'];
          for (const name of possibleNames) {
            if (window[name] && typeof window[name].renderAsync === 'function') {
              docxLib = window[name];
              break;
            }
          }
        }
        
        if (!docxLib || typeof docxLib.renderAsync !== 'function') {
          throw new Error('docx-preview renderAsync method not found. Available globals: ' + Object.keys(window).filter(k => k.toLowerCase().includes('docx')).join(', '));
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
          useBase64URL: true,  // Convert images to base64 URLs
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
                img.onerror = resolve; // Continue even if image fails
                setTimeout(resolve, 3000); // Timeout after 3 seconds
              }
            });
          }));
        }

        // Signal completion
        window.conversionComplete = true;
        window.conversionError = null;
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

    // Extract the rendered HTML
    const htmlContent = await page.evaluate(() => {
      const container = document.getElementById('container');
      if (!container) {
        throw new Error('Container not found');
      }
      
      // Get the rendered content (docx-preview wraps it in .docx-wrapper)
      const wrapper = container.querySelector('.docx-wrapper') || container;
      return wrapper.innerHTML;
    });

    if (!htmlContent || htmlContent.trim().length === 0) {
      throw new Error("Rendered HTML is empty");
    }

    // Wrap in proper HTML document structure
    const finalHtml = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      margin: 0;
      padding: 20px;
      background: white;
      font-family: 'Calibri', 'Arial', 'Helvetica', sans-serif;
    }
    /* Preserve all styles from docx-preview */
    .docx-wrapper {
      background: white !important;
      padding: 0 !important;
    }
    /* Ensure all inline styles from Word are preserved */
    * {
      box-sizing: border-box;
    }
  </style>
</head>
<body>
  <div class="docx-wrapper">
    ${htmlContent}
  </div>
</body>
</html>`;

    console.log("Word to HTML conversion successful using docx-preview. HTML length:", finalHtml.length, "characters");

    // Close browser
    await browser.close();
    browser = null;

    if (format === "json") {
      // Return HTML as JSON
      const response = {
        success: true,
        html: finalHtml
      };

      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      return res.json(response);
    } else {
      // Return HTML directly
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      return res.send(finalHtml);
    }
  } catch (error) {
    console.error("Word to HTML conversion error:", error);
    
    // Ensure browser is closed on error
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
