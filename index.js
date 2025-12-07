import express from "express";
import cors from "cors";
import puppeteer from "puppeteer";

const app = express();
app.use(cors());
app.use(express.json({ limit: "10mb" }));

// API key security
app.use((req, res, next) => {
    const apiKey = process.env.API_KEY;
    if (apiKey && req.headers["x-api-key"] !== apiKey) {
        return res.status(403).json({ error: "Unauthorized" });
    }
    next();
});

app.post("/convert", async (req, res) => {
    try {
        const { html } = req.body;

        if (!html) {
            return res.status(400).json({ error: "HTML content required." });
        }

        const browser = await puppeteer.launch({
            headless: true,
            args: [
              "--no-sandbox",
              "--disable-setuid-sandbox",
              "--disable-dev-shm-usage",
              "--disable-gpu"
            ],
            executablePath: process.env.PUPPETEER_EXECUTABLE_PATH
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

    } catch (err) {
        console.error("PDF ERROR:", err);
        return res.status(500).json({
            success: false,
            error: "Failed to generate PDF",
            details: err.message,
        });
    }
});

app.get("/", (req, res) => {
    res.send("HTML to PDF Service Running ✔");
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on ${PORT}`));
