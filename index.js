import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import chromium from "@sparticuz/chromium";
import puppeteer from "puppeteer-core";

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: "50mb" })); // Increased limit

// Use ?format=json to get base64 for Salesforce, else download PDF directly
app.post("/convert", async (req, res) => {
  try {
    const { html } = req.body;
    const format = req.query.format?.toLowerCase() === "json" ? "json" : "pdf";

    if (!html) {
      return res.status(400).json({
        success: false,
        error: "HTML content is required"
      });
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
      // Sanitize Base64 (remove newlines) before sending to Salesforce
      const base64Pdf = pdfBuffer.toString("base64").replace(/\r?\n|\r/g, "");
      return res.json({
        success: true,
        pdf: base64Pdf
      });
    } else {
      // Raw PDF download
      res.set({
        "Content-Type": "application/pdf",
        "Content-Length": pdfBuffer.length
      });
      return res.send(pdfBuffer);
    }

  } catch (error) {
    console.error("PDF error:", error);
    return res.status(500).json({
      success: false,
      error: "Failed to generate PDF",
      details: error.message
    });
  }
});

app.get("/", (req, res) => {
  res.send("HTML to PDF Service is running");
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on ${PORT}`));
