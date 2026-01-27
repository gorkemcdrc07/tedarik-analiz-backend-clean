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

// küçük yardımcı: sleep
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// fetch + timeout + retry (undici YOK)
async function fetchWithTimeoutAndRetry(
    url,
    options,
    { timeoutMs = 120_000, retries = 2, retryDelayMs = 1000 } = {}
) {
    let lastErr;

    for (let attempt = 0; attempt <= retries; attempt++) {
        const controller = new AbortController();
        const t = setTimeout(() => controller.abort(), timeoutMs);

        try {
            const res = await fetch(url, {
                ...options,
                signal: controller.signal,
            });
            clearTimeout(t);
            return res;
        } catch (err) {
            clearTimeout(t);
            lastErr = err;

            const msg = String(err?.message || "").toLowerCase();
            const retryable =
                msg.includes("fetch failed") ||
                msg.includes("abort") ||
                msg.includes("timeout") ||
                msg.includes("network");

            if (attempt < retries && retryable) {
                await sleep(retryDelayMs * (attempt + 1));
                continue;
            }

            throw err;
        }
    }

    throw lastErr;
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

        const upstreamUrl =
            "https://api.odaklojistik.com.tr/api/tmsorders/getall";

        const upstreamRes = await fetchWithTimeoutAndRetry(
            upstreamUrl,
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    Authorization: token, // gerekirse: `Bearer ${token}`
                },
                body: JSON.stringify({ startDate, endDate, userId }),
            },
            {
                timeoutMs: 120_000,
                retries: 2,
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

        return res.status(500).json({
            rid,
            error: "Backend exception",
            message: err?.message || "fetch failed",
        });
    }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () =>
    console.log("🚀 Server listening on port", PORT)
);
