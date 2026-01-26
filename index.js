const express = require("express");
const cors = require("cors");

// ✅ node-fetch v3 (CommonJS) uyumlu fetch
const fetch = (...args) =>
  import("node-fetch").then(({ default: fetch }) => fetch(...args));

const app = express();

app.use(cors());
app.use(express.json({ limit: "5mb" }));

app.get("/", (req, res) => {
  res.send("Tedarik Analiz Backend is running");
});

app.get("/health", (req, res) => {
  res.json({ ok: true, service: "tedarik-analiz-backend" });
});

app.get("/times-by-sefer", (req, res) => {
  res.json({
    "SEFER-001": {
      yukleme_varis: "2026-01-01T08:00:00.000Z",
      yukleme_giris: "2026-01-01T08:10:00.000Z",
      yukleme_cikis: "2026-01-01T08:30:00.000Z",
      teslim_varis: "2026-01-01T12:00:00.000Z",
      teslim_giris: "2026-01-01T12:05:00.000Z",
      teslim_cikis: "2026-01-01T12:20:00.000Z",
    },
  });
});

// 🔥 ODAK API PROXY (DEBUG’LI)
app.post("/tmsorders", async (req, res) => {
  const rid = `${Date.now()}-${Math.random().toString(16).slice(2)}`;

  try {
    const { startDate, endDate, userId } = req.body || {};
    console.log(`[${rid}] /tmsorders body:`, req.body);

    if (!startDate || !endDate || userId == null) {
      return res.status(400).json({
        rid,
        error: "startDate, endDate ve userId zorunlu",
      });
    }

    const token = process.env.ODAK_API_TOKEN;
    if (!token) {
      return res.status(500).json({ rid, error: "ODAK_API_TOKEN env yok" });
    }

    const upstreamUrl = "https://api.odaklojistik.com.tr/api/tmsorders/getall";

    // önce raw Authorization deniyoruz; gerekirse Bearer’a çevireceğiz
    const headers = {
      "Content-Type": "application/json",
      Authorization: token,
      // Authorization: `Bearer ${token}`, // gerekirse buna çevireceğiz
    };

    const upstreamRes = await fetch(upstreamUrl, {
      method: "POST",
      headers,
      body: JSON.stringify({ startDate, endDate, userId }),
    });

    const rawText = await upstreamRes.text();
    console.log(`[${rid}] upstream status:`, upstreamRes.status);
    console.log(`[${rid}] upstream text (first 500):`, rawText.slice(0, 500));

    let data = null;
    try {
      data = rawText ? JSON.parse(rawText) : null;
    } catch (e) {
      return res.status(502).json({
        rid,
        error: "Upstream JSON değil / parse edilemedi",
        upstreamStatus: upstreamRes.status,
        upstreamText: rawText.slice(0, 1000),
      });
    }

    if (!upstreamRes.ok) {
      return res.status(502).json({
        rid,
        error: "Upstream hata döndürdü",
        upstreamStatus: upstreamRes.status,
        upstreamData: data,
      });
    }

    return res.json({ rid, ok: true, data });
  } catch (err) {
    console.error(`[${rid}] /tmsorders exception:`, err);
    return res.status(500).json({
      rid,
      error: "Backend exception",
      message: err?.message,
    });
  }
});

app.post("/analyze", (req, res) => {
  res.json({ ok: true, received: req.body });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server listening on port", PORT));
