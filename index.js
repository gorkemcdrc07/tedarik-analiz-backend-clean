// index.js (backend) — GÜNCEL TAM KOD (undici yok, fail-fast timeout var, optional cache var)

const express = require("express");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json({ limit: "2mb" }));

// Root
app.get("/", (req, res) => {
    res.send("Tedarik Analiz Backend is running");
});

// Health
app.get("/health", (req, res) => {
    res.json({ ok: true, service: "tedarik-analiz-backend-clean" });
});

// küçük yardımcı
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ✅ basit memory cache (aynı sorguyu tekrar tekrar Odak'a atmasın)
const cache = new Map(); // key -> { ts, value }
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 dk

function cacheKey({ startDate, endDate, userId }) {
    return `${userId}|${startDate}|${endDate}`;
}

async function fetchOdakWithTimeout(url, options, timeoutMs) {
    const controller = new AbortController();
    const t = setTimeout(() => controller.abort(), timeoutMs);
    try {
        return await fetch(url, { ...options, signal: controller.signal });
    } finally {
        clearTimeout(t);
    }
}

// 🔥 ODAK API PROXY
app.post("/tmsorders", async (req, res) => {
    const rid = `${Date.now()}-${Math.random().toString(16).slice(2)}`;

    try {
        const { startDate, endDate, userId } = req.body || {};

        if (!startDate || !endDate || userId == null) {
            return res.status(400).json({
                rid,
                error: "startDate, endDate ve userId zorunlu",
            });
        }

        const token = process.env.ODAK_API_TOKEN;
        if (!token) {
            return res.status(500).json({
                rid,
                error: "ODAK_API_TOKEN Render env'de yok",
            });
        }

        // ✅ Cache kontrol
        const key = cacheKey({ startDate, endDate, userId });
        const hit = cache.get(key);
        if (hit && Date.now() - hit.ts < CACHE_TTL_MS) {
            return res.json({ rid, ok: true, data: hit.value, cached: true });
        }

        const upstreamUrl = "https://api.odaklojistik.com.tr/api/tmsorders/getall";

        // ⚠️ Token formatı: Odak Bearer istiyorsa bu doğru.
        // Eğer Odak direkt token istiyorsa alt satırı `Authorization: token` yap.
        const authHeader = token.startsWith("Bearer ") ? token : `Bearer ${token}`;

        const options = {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: authHeader,
            },
            body: JSON.stringify({ startDate, endDate, userId }),
        };

        // ✅ Fail-fast timeout: Render gateway 502 basmadan biz kontrollü hata dönelim
        const TIMEOUT_MS = 25_000; // 25 sn
        const RETRIES = 1; // 1 retry yeterli (toplam max ~50 sn)

        let upstreamRes = null;
        let lastErr = null;

        for (let attempt = 0; attempt <= RETRIES; attempt++) {
            try {
                upstreamRes = await fetchOdakWithTimeout(upstreamUrl, options, TIMEOUT_MS);
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
                error: "Odak JSON dönmedi",
                status: upstreamRes.status,
                raw: text.slice(0, 800),
            });
        }

        if (!upstreamRes.ok) {
            return res.status(502).json({
                rid,
                error: "Odak API hata döndü",
                status: upstreamRes.status,
                data,
            });
        }

        // ✅ cache yaz
        cache.set(key, { ts: Date.now(), value: data });

        return res.json({ rid, ok: true, data });
    } catch (err) {
        console.error("❌ ODAK API HATASI:", err);
        return res.status(500).json({
            rid,
            error: "Backend exception",
            message: err?.message || "fetch failed",
        });
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Server listening on port", PORT));
