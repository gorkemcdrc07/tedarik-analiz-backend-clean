const express = require("express");
const cors = require("cors");

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

// Frontend'in beklediği formatta örnek veri
app.get("/times-by-sefer", (req, res) => {
  res.json({
    "SEFER-001": {
      yukleme_varis: "2026-01-01T08:00:00.000Z",
      yukleme_giris: "2026-01-01T08:10:00.000Z",
      yukleme_cikis: "2026-01-01T08:30:00.000Z",
      teslim_varis: "2026-01-01T12:00:00.000Z",
      teslim_giris: "2026-01-01T12:05:00.000Z",
      teslim_cikis: "2026-01-01T12:20:00.000Z"
    },
    "SEFER-002": {
      yukleme_varis: "2026-01-02T09:00:00.000Z",
      yukleme_giris: null,
      yukleme_cikis: null,
      teslim_varis: "2026-01-02T13:30:00.000Z",
      teslim_giris: null,
      teslim_cikis: null
    }
  });
});

// (opsiyonel) POST test
app.post("/analyze", (req, res) => {
  res.json({ ok: true, received: req.body });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server listening on port", PORT));
