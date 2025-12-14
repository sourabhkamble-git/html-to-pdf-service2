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

// New endpoint for Word to HTML conversion using mammoth.js
// mammoth.js preserves merge fields ({{FieldName}}) AND styling
// This is the best approach for server-side automation where merge fields are critical
app.post("/convert-word-to-html", async (req, res) => {
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

    console.log("Converting Word to HTML using mammoth.js (preserves merge fields + styling). File size:", wordBuffer.length, "bytes");

    // Enhanced style mapping to preserve Word formatting
    const styleMap = [
      "p[style-name='Title'] => h1.title:fresh",
      "p[style-name='Heading 1'] => h1:fresh",
      "p[style-name='Heading 2'] => h2:fresh",
      "p[style-name='Heading 3'] => h3:fresh",
      "r[style-name='Strong'] => strong",
      "r[style-name='Emphasis'] => em",
      "p[style-name='Normal'] => p"
    ];

    // Convert Word document to HTML using mammoth
    // mammoth preserves merge fields as {{FieldName}} AND inline styles (colors, fonts, alignment)
    const result = await mammoth.convertToHtml(
      { buffer: wordBuffer },
      {
        styleMap: styleMap,
        includeDefaultStyleMap: true, // Include default style mappings
        includeEmbeddedStyleMap: true // Include styles embedded in the Word document
      }
    );

    let htmlContent = result.value; // The HTML content with merge fields preserved
    const messages = result.messages; // Any warnings or errors

    // Log any conversion messages
    if (messages && messages.length > 0) {
      console.log("Mammoth conversion messages:", messages);
    }

    if (!htmlContent || htmlContent.trim().length === 0) {
      return res.status(500).json({ 
        success: false, 
        error: "Word to HTML conversion produced empty HTML content" 
      });
    }

    // Check if merge fields are preserved
    const mergeFieldPattern = /\{\{[^}]+\}\}/g;
    const mergeFields = htmlContent.match(mergeFieldPattern) || [];
    console.log("Merge fields preserved in HTML:", mergeFields.length, "fields");
    if (mergeFields.length > 0) {
      console.log("Sample merge fields:", mergeFields.slice(0, 5));
    }

    // Wrap HTML content in a proper document structure with enhanced CSS
    // This ensures styles, colors, fonts, and alignment are preserved
    htmlContent = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    /* Base styles - mammoth.js outputs inline styles, so we preserve them */
    body {
      font-family: 'Calibri', 'Arial', 'Helvetica', sans-serif;
      margin: 20px;
      color: #000000;
      line-height: 1.6;
    }
    
    /* Preserve table formatting */
    table {
      border-collapse: collapse;
      width: 100%;
      margin: 10px 0;
      border-spacing: 0;
    }
    
    td, th {
      border: 1px solid #ddd;
      padding: 8px;
      text-align: left;
      vertical-align: top;
    }
    
    th {
      background-color: #f2f2f2;
      font-weight: bold;
    }
    
    /* Ensure all inline styles from Word are preserved */
    * {
      box-sizing: border-box;
    }
    
    /* Preserve merge fields styling - make them visible */
    /* Merge fields like {{Opportunity.Name}} will appear as-is in the HTML */
  </style>
</head>
<body>
${htmlContent}
</body>
</html>`;

    console.log("Word to HTML conversion successful. HTML length:", htmlContent.length, "characters");
    console.log("Merge fields count:", mergeFields.length);

    if (format === "json") {
      // Return HTML as JSON
      const response = {
        success: true,
        html: htmlContent
      };

      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      return res.json(response);
    } else {
      // Return HTML directly
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      return res.send(htmlContent);
    }
  } catch (error) {
    console.error("Word to HTML conversion error:", error);
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
