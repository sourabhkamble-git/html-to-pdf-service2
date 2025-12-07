import express from "express";
import cors from "cors";
import bodyParser from "body-parser";
import chromium from "@sparticuz/chromium";
import puppeteer from "puppeteer-core";

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: "20mb" }));

app.post("/generate-pdf", async (req, res) => {
  try {
    const { html } = req.body;

    if (!html) {
      return res.status(400).json({
        success: false,
        error: "HTML content is required"
      });
    }

    // Launch browser using chromium + puppeteer-core
    const browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: chromium.defaultViewport,
      executablePath: await chromium.executablePath(),
      headless: true
    });

    const page = await browser.newPage();

    await page.setContent(html, {
      waitUntil: ["load", "domcontentloaded"]
    });

    const pdfBuffer = await page.pdf({
      format: "A4",
      printBackground: true
    });

    await browser.close();

    res.set({
      "Content-Type": "application/pdf",
      "Content-Length": pdfBuffer.length
    });

    res.send(pdfBuffer);

  } catch (error) {
    console.error("PDF error:", error);

    res.status(500).json({
      success: false,
      error: "Failed to generate PDF",
      details: error.message
    });
  }
});

app.get("/", (req, res) => {
  res.send("HTML to PDF Service is running");
});

// Render uses PORT automatically
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`Server running on ${PORT}`));
