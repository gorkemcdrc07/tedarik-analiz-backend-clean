// index.js (backend) - GÜNCEL TAM KOD
const express = require("express");
const cors = require("cors");

// ✅ undici timeouts (Node fetch altyapısı)
const { setGlobalDispatcher, Agent } = require("undici");

// Global fetch timeoutları artır (Render + yavaş upstream için şart)
setGlobalDispatcher(
    new Agent({
        connectTimeout: 30_000,
        headersTimeout: 120_000, // ✅ en kritik
        bodyTimeout: 120_000,
    })
);

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

// küçük yardımcı: sleep
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// küçük yardımcı: fetch + timeout + retry
async function fetchWithTimeoutAndRetry(url, options, { timeoutMs = 120_000, retries = 2, retryDelayMs = 800 } = {}) {
    let lastErr = null;

    for (let attempt = 0; attempt <= retries; attempt++) {
        const controller = new AbortController();
        const t = setTimeout(() => controller.abort(), timeoutMs);

        try {
            const resp = await fetch(url, {
                ...options,
                signal: controller.signal,
            });
            clearTimeout(t);
            return resp;
        } catch (err) {
            clearTimeout(t);
            lastErr = err;

            // AbortError / timeout / network hatalarında retry yap
            const msg = String(err?.message || "");
            const code = err?.cause?.code || err?.code;

            const retryable =
                msg.toLowerCase().includes("fetch failed") ||
                msg.toLowerCase().includes("abort") ||
                code === "UND_ERR_HEADERS_TIMEOUT" ||
                code === "UND_ERR_CONNECT_TIMEOUT" ||
                code === "ETIMEDOUT" ||
                code === "ECONNRESET" ||
                code === "ECONNREFUSED" ||
                code === "ENOTFOUND";

            if (attempt < retries && retryable) {
                await sleep(retryDelayMs * (attempt + 1));
                continue;
            }

            throw err;
        }
    }

    // buraya normalde düşmez
    throw lastErr || new Error("fetch failed");
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

        const upstreamUrl = "https://api.odaklojistik.com.tr/api/tmsorders/getall";

        // ✅ Authorization formatı önemli olabilir:
        // Eğer Odak "Bearer <token>" istiyorsa aşağıdaki satırı aç, üsttekini kapat.
        const authHeader = token.startsWith("Bearer ") ? token : token; // default: direkt token
        // const authHeader = token.startsWith("Bearer ") ? token : `Bearer ${token}`;

        const upstreamRes = await fetchWithTimeoutAndRetry(
            upstreamUrl,
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    Authorization: authHeader,
                },
                body: JSON.stringify({ startDate, endDate, userId }),
            },
            {
                timeoutMs: 120_000, // Odak yavaşsa artır: 180000
                retries: 2,
                retryDelayMs: 900,
            }
        );

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

        return res.json({ rid, ok: true, data });
    } catch (err) {
        console.error("❌ ODAK API HATASI:", err);

        // daha açıklayıcı hata (timeout code vs.)
        const code = err?.cause?.code || err?.code || null;

        return res.status(500).json({
            rid,
            error: "Backend exception",
            message: err?.message || "fetch failed",
            code,
        });
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log("🚀 Server listening on port", PORT));
