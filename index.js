// index.js — GÜNCEL TAM KOD (Express 5 uyumlu CORS preflight fix + /tmsorders + /tmsorders/week)

const express = require("express");
const cors = require("cors");

const app = express();

/* =======================
   ✅ CORS (PRE-FLIGHT FIX)
   ======================= */

const ALLOWED_ORIGINS = [
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "https://analiz-pearl.vercel.app",
    "https://analiz-v2.vercel.app", // 👈 BUNU EKLE
];

const corsOptions = {
    origin: (origin, cb) => {
        // origin yoksa (Postman / server-side) izin ver
        if (!origin) return cb(null, true);

        // allowlist
        if (ALLOWED_ORIGINS.includes(origin)) return cb(null, true);

        // ❗ error fırlatma -> bazı ortamlarda header basılmadan kapanır
        return cb(null, false);
    },
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: false,
    maxAge: 86400,
};

// ✅ önce CORS
app.use(cors(corsOptions));

// ✅ Express 5 + path-to-regexp uyumlu preflight ( "*" KULLANMA! )
app.options(/.*/, cors(corsOptions));

// JSON body
app.use(express.json({ limit: "2mb" }));

/* =======================
   ROUTES
   ======================= */

app.get("/", (req, res) => {
    res.send("Tedarik Analiz Backend is running");
});

app.get("/health", (req, res) => {
    res.json({ ok: true, service: "tedarik-analiz-backend-clean" });
});

app.get("/routes", (req, res) => {
    res.json({
        routes: [
            "GET /",
            "GET /health",
            "GET /routes",
            "POST /tmsorders",
            "POST /tmsorders/week",
        ],
        allowedOrigins: ALLOWED_ORIGINS,
    });
});

/* =======================
   Helpers
   ======================= */

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const cache = new Map(); // key -> { ts, value }
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 dk

function cacheKey({ startDate, endDate, userId }) {
    return `${userId}|${startDate}|${endDate}`;
}

async function fetchWithTimeout(url, options, timeoutMs) {
    const controller = new AbortController();
    const t = setTimeout(() => controller.abort(), timeoutMs);
    try {
        return await fetch(url, { ...options, signal: controller.signal });
    } finally {
        clearTimeout(t);
    }
}

/* =======================
   🔥 ODAK API PROXY
   ======================= */

async function tmsordersHandler(req, res) {
    const rid = `${Date.now()}-${Math.random().toString(16).slice(2)}`;

    try {
        const { startDate, endDate, userId } = req.body || {};

        if (!startDate || !endDate || userId == null) {
            return res.status(400).json({
                rid,
                ok: false,
                error: "startDate, endDate ve userId zorunlu",
            });
        }

        const token = process.env.ODAK_API_TOKEN;
        if (!token) {
            return res.status(500).json({
                rid,
                ok: false,
                error: "ODAK_API_TOKEN Render env'de yok",
            });
        }

        // ✅ Cache hit
        const key = cacheKey({ startDate, endDate, userId });
        const hit = cache.get(key);
        if (hit && Date.now() - hit.ts < CACHE_TTL_MS) {
            return res.json({ rid, ok: true, cached: true, data: hit.value });
        }

        const upstreamUrl = "https://api.odaklojistik.com.tr/api/tmsorders/getall";

        // Odak token formatı: Bearer gerekiyorsa böyle, değilse token direkt gönder
        const authHeader = token.startsWith("Bearer ") ? token : `Bearer ${token}`;

        const options = {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: authHeader,
            },
            body: JSON.stringify({ startDate, endDate, userId }),
        };

        const TIMEOUT_MS = 25_000;
        const RETRIES = 1;

        let upstreamRes = null;
        let lastErr = null;

        for (let attempt = 0; attempt <= RETRIES; attempt++) {
            try {
                upstreamRes = await fetchWithTimeout(upstreamUrl, options, TIMEOUT_MS);
                lastErr = null;
                break;
            } catch (e) {
                lastErr = e;
                if (attempt < RETRIES) await sleep(600);
            }
        }

        if (!upstreamRes) {
            return res.status(504).json({
                rid,
                ok: false,
                error: "Odak timeout / erişilemiyor",
                message: lastErr?.message || "fetch failed",
            });
        }

        const text = await upstreamRes.text();

        let data;
        try {
            data = text ? JSON.parse(text) : null;
        } catch {
            return res.status(502).json({
                rid,
                ok: false,
                error: "Odak JSON dönmedi",
                status: upstreamRes.status,
                raw: text.slice(0, 800),
            });
        }

        if (!upstreamRes.ok) {
            return res.status(502).json({
                rid,
                ok: false,
                error: "Odak API hata döndü",
                status: upstreamRes.status,
                data,
            });
        }

        cache.set(key, { ts: Date.now(), value: data });

        return res.json({ rid, ok: true, cached: false, data });
    } catch (err) {
        console.error("❌ BACKEND EXCEPTION:", err);
        return res.status(500).json({
            rid: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
            ok: false,
            error: "Backend exception",
            message: err?.message || String(err),
        });
    }
}

// ✅ iki endpoint de aynı handler
app.post("/tmsorders", tmsordersHandler);
app.post("/tmsorders/week", tmsordersHandler);

/* =======================
   START
   ======================= */

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Server listening on port", PORT));
