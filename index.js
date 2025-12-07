import express from "express";
import cors from "cors";
import puppeteer from "puppeteer";

const app = express();
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// Simple API key auth (optional but recommended)
app.use((req, res, next) => {
    const apiKey = process.env.API_KEY;
    if (!apiKey) return next(); // No key set = no auth

    if (req.headers["x-api-key"] !== apiKey) {
        return res.status(403).json({ error: "Unauthorized" });
    }
    next();
});

app.post("/convert", async (req, res) => {
    try {
        const { html } = req.body;

        if (!html || html.trim() === "") {
            return res.status(400).json({ error: "HTML content is required." });
        }

        const browser = await puppeteer.launch({
            headless: "new",
            args: ["--no-sandbox", "--disable-setuid-sandbox"],
        });

        const page = await browser.newPage();
        await page.setContent(html, { waitUntil: "networkidle0" });

        const pdfBuffer = await page.pdf({
            format: "A4",
            printBackground: true,
        });

        await browser.close();

        return res.json({
            success: true,
            pdf: pdfBuffer.toString("base64"),
        });
    } catch (error) {
        console.error("PDF generation error:", error);
        return res.status(500).json({
            success: false,
            error: "Failed to generate PDF",
            details: error.message,
        });
    }
});

app.get("/", (req, res) => {
    res.send("HTML to PDF Service is Running ✔");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`PDF service running on port ${PORT}`));
