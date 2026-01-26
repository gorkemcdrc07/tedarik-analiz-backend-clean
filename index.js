const express = require("express");
const cors = require("cors");
const fetch = require("node-fetch");

const app = express();

app.use(cors());
app.use(express.json());

// Root
app.get("/", (req, res) => {
    res.send("Tedarik Analiz Backend is running");
});

// Health check
app.get("/health", (req, res) => {
    res.json({ ok: true, service: "tedarik-analiz-backend" });
});

// Frontend'in beklediği formatta örnek veri (ŞİMDİLİK DURUYOR)
app.get("/times-by-sefer", (req, res) => {
    res.json({
        "SEFER-001": {
            yukleme_varis: "2026-01-01T08:00:00.000Z",
            yukleme_giris: "2026-01-01T08:10:00.000Z",
            yukleme_cikis: "2026-01-01T08:30:00.000Z",
            teslim_varis: "2026-01-01T12:00:00.000Z",
            teslim_giris: "2026-01-01T12:05:00.000Z",
            teslim_cikis: "2026-01-01T12:20:00.000Z"
        }
    });
});

// 🔥 GERÇEK ODAK API PROXY ENDPOINT
app.post("/tmsorders", async (req, res) => {
    try {
        const { startDate, endDate, userId } = req.body;

        if (!startDate || !endDate || !userId) {
            return res.status(400).json({
                error: "startDate, endDate ve userId zorunlu"
            });
        }

        const response = await fetch(
            "https://api.odaklojistik.com.tr/api/tmsorders/getall",
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": process.env.ODAK_API_TOKEN
                },
                body: JSON.stringify({
                    startDate,
                    endDate,
                    userId
                })
            }
        );

        const data = await response.json();
        res.json(data);

    } catch (err) {
        console.error("❌ ODAK API HATASI:", err);
        res.status(500).json({ error: "Odak API çağrısı başarısız" });
    }
});

// (opsiyonel) POST test
app.post("/analyze", (req, res) => {
    res.json({ ok: true, received: req.body });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server listening on port", PORT));
