// index.js

const express = require("express");
const cors = require("cors");
const dotenv = require("dotenv");
const puppeteer = require("puppeteer");
const cron = require("node-cron");
const fs = require("fs");
const path = require("path");
const dns = require("dns");
const XLSX = require("xlsx-js-style");

dns.setDefaultResultOrder("ipv4first");
dotenv.config();

const app = express();
const PORT = process.env.PORT || 10000;

const ALLOWED_ORIGINS = [
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "http://localhost:3001",
    "http://127.0.0.1:3001",
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "https://analiz-pearl.vercel.app",
    "https://analiz-v2.vercel.app",
];

const corsOptions = {
    origin: (origin, cb) => {
        if (!origin) return cb(null, true);
        if (ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
        console.warn("⛔ CORS blocked origin:", origin);
        return cb(null, false);
    },
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: false,
    maxAge: 86400,
};

app.use(cors(corsOptions));
app.options(/.*/, cors(corsOptions));
app.use(express.json({ limit: "10mb" }));

/* ======================= MAIL AYARLARI ======================= */

const sendMailWithResend = async ({ to, cc, subject, text, attachments }) => {
    const body = {
        from: "Odak Lojistik <onboarding@resend.dev>",
        to: [to],
        subject,
        text,
    };

    if (cc && cc.length > 0) body.cc = cc;

    if (attachments && attachments.length > 0) {
        body.attachments = attachments.map((a) => ({
            filename: a.filename,
            content: Buffer.from(a.content).toString("base64"),
        }));
    }

    const res = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            Authorization: `Bearer ${process.env.RESEND_API_KEY}`,
        },
        body: JSON.stringify(body),
    });

    if (!res.ok) {
        const err = await res.text();
        throw new Error(`Resend hata: ${err}`);
    }

    return await res.json();
};

const LAST_PAYLOAD_FILE = path.join(__dirname, "last-mail-payload.json");

const sonPayloadKaydet = (payload) => {
    fs.writeFileSync(LAST_PAYLOAD_FILE, JSON.stringify(payload, null, 2), "utf8");
};

const sonPayloadOku = () => {
    if (!fs.existsSync(LAST_PAYLOAD_FILE)) return null;
    return JSON.parse(fs.readFileSync(LAST_PAYLOAD_FILE, "utf8"));
};

const escapeHtml = (v) =>
    String(v ?? "-")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");

const slugify = (t) =>
    String(t || "rapor")
        .toLowerCase()
        .trim()
        .replace(/ğ/g, "g")
        .replace(/ü/g, "u")
        .replace(/ş/g, "s")
        .replace(/ı/g, "i")
        .replace(/ö/g, "o")
        .replace(/ç/g, "c")
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "");

const pick = (...vals) =>
    vals.find((v) => v !== undefined && v !== null && String(v).trim() !== "");

const getYukleme = (r = {}) =>
    pick(
        r.PickupAddressCode,
        r.yuklemeNoktasi,
        r.yuklemeNokta,
        r.yuklemeNoktaAdi,
        r.yuklemeYeri,
        r.yuklemeAdres,
        r.yuklemeLokasyon,
        r.cikisNoktasi,
        r.cikisYeri,
        r.gondericiUnvan,
        r.gonderici,
        r.gonderen,
        "-"
    );

const getTeslim = (r = {}) =>
    pick(
        r.DeliveryAddressCode,
        r.teslimNoktasi,
        r.teslimNokta,
        r.teslimNoktaAdi,
        r.teslimYeri,
        r.teslimAdres,
        r.teslimLokasyon,
        r.varisNoktasi,
        r.varisYeri,
        r.aliciUnvan,
        r.alici,
        "-"
    );

const formatNokta = (...vals) => {
    return (
        vals
            .filter(
                (v) =>
                    v !== undefined &&
                    v !== null &&
                    String(v).trim() !== ""
            )
            .map((v) => String(v).trim())
            .filter((v) => v !== "-")
            .join(" / ") || "-"
    );
};
const formatDateLong = (v) => {
    const d = v ? new Date(v) : new Date();
    return isNaN(d.getTime())
        ? "-"
        : d.toLocaleDateString("tr-TR", {
            day: "2-digit",
            month: "long",
            year: "numeric",
        });
};

const short = (t, max = 42) => {
    const s = String(t || "-");
    return s.length > max ? s.slice(0, max) + "..." : s;
};

const statusText = (r) => {
    if (!r?.yuklemeyeGelis || String(r.yuklemeyeGelis).trim() === "-") {
        return "Yükleme Tarihi Yok";
    }
    return r?.durumText || "-";
};

const statusClass = (r) => {
    const s = statusText(r);
    if (s === "Zamanında") return "bs";
    if (s === "Geç Tedarik") return "bd";
    if (s === "Yükleme Tarihi Yok") return "bw";
    return "bn";
};

const hesaplaOzet = (summary = {}, rows = []) => {
    return {
        ...summary,
        talep: Number(summary.talep || 0),
        tedarik: Number(summary.tedarik || 0),
        edilmeyen: Number(summary.edilmeyen || 0),
        gec_tedarik: Number(summary.gec_tedarik || 0),
        sho_basilan: Number(summary.sho_basilan || 0),
        sho_basilmayan: Number(summary.sho_basilmayan || 0),
        spot: Number(summary.spot || 0),
        filo: Number(summary.filo || 0),
    };
};

const perfStyle = (p) => {
    if (p >= 90) return { bg: "#f0fdf4", border: "#bbf7d0", color: "#166534", bar: "#16a34a" };
    if (p >= 70) return { bg: "#dbeafe", border: "#bfdbfe", color: "#1d4ed8", bar: "#2563eb" };
    if (p >= 50) return { bg: "#fef3c7", border: "#fde68a", color: "#92400e", bar: "#f59e0b" };
    return { bg: "#fee2e2", border: "#fecaca", color: "#991b1b", bar: "#ef4444" };
};

const DOT_COLORS = [
    "#818cf8", "#f59e0b", "#34d399", "#f87171",
    "#60a5fa", "#a78bfa", "#fb7185", "#38bdf8",
];

const CSS = `
@page { size: A4 landscape; margin: 8mm 6mm; }
* { box-sizing: border-box; }
body { margin: 0; background: #eef3f8; font-family: Arial, Helvetica, sans-serif; color: #0f172a; }
.wrap { padding: 8px; }
.hero { background: #0f172a; border-radius: 16px; padding: 18px 22px; margin-bottom: 13px; display: flex; justify-content: space-between; align-items: flex-end; }
.hb { font-size: 9px; font-weight: 800; letter-spacing: .16em; text-transform: uppercase; color: #93c5fd; margin-bottom: 4px; }
.ht { font-size: 22px; font-weight: 900; color: #fff; letter-spacing: -.03em; }
.hs { font-size: 10px; color: #64748b; margin-top: 3px; }
.hd { font-size: 10px; color: #94a3b8; text-align: right; line-height: 1.7; }
.gkpi { background: #fff; border: 1px solid #e2e8f0; border-radius: 13px; padding: 12px 14px; margin-bottom: 13px; }
.gkpi-h { font-size: 8px; font-weight: 800; letter-spacing: .1em; text-transform: uppercase; color: #64748b; margin-bottom: 9px; }
.krow { display: flex; gap: 7px; flex-wrap: wrap; }
.kp { flex: 1; min-width: 78px; border-radius: 9px; padding: 9px 10px; text-align: center; }
.kp .l { font-size: 7px; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; margin-bottom: 4px; }
.kp .v { font-size: 20px; font-weight: 900; line-height: 1; letter-spacing: -.03em; }
.kp .s { font-size: 7.5px; margin-top: 3px; }
.kp-b { background: #eff6ff; border: 1px solid #bfdbfe; }
.kp-b .l, .kp-b .v { color: #1d4ed8; }
.kp-g { background: #f0fdf4; border: 1px solid #bbf7d0; }
.kp-g .l, .kp-g .v { color: #166534; }
.kp-r { background: #fff1f2; border: 1px solid #fecdd3; }
.kp-r .l, .kp-r .v { color: #991b1b; }
.kp-a { background: #fff7ed; border: 1px solid #fed7aa; }
.kp-a .l, .kp-a .v { color: #9a3412; }
.kp-p { background: #faf5ff; border: 1px solid #e9d5ff; }
.kp-p .l, .kp-p .v { color: #6b21a8; }
.kp-n { background: #f8fafc; border: 1px solid #e2e8f0; }
.kp-n .l, .kp-n .v { color: #94a3b8; }
.pkp { flex: 1.3; min-width: 88px; border-radius: 9px; padding: 9px 12px; text-align: center; }
.pkp .l { font-size: 7px; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; margin-bottom: 4px; }
.pkp .v { font-size: 20px; font-weight: 900; line-height: 1; }
.bw2 { margin-top: 5px; height: 4px; background: rgba(0,0,0,.09); border-radius: 2px; overflow: hidden; }
.br { height: 100%; border-radius: 2px; }
.pdiv { display: flex; align-items: center; gap: 9px; margin: 15px 0 9px; }
.pdiv-line { flex: 1; height: 1px; background: #e2e8f0; }
.pdiv-badge { display: flex; align-items: center; gap: 5px; background: #0f172a; color: #e2e8f0; border-radius: 7px; padding: 4px 11px; font-size: 8.5px; font-weight: 800; letter-spacing: .06em; text-transform: uppercase; white-space: nowrap; }
.pdiv-dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
.pozet { background: #fff; border: 1px solid #e2e8f0; border-radius: 11px; padding: 10px 12px; margin-bottom: 9px; display: flex; gap: 7px; align-items: stretch; }
.ps { flex: 1; border-radius: 7px; padding: 7px 9px; text-align: center; }
.ps .l { font-size: 7px; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; margin-bottom: 3px; }
.ps .v { font-size: 16px; font-weight: 900; line-height: 1; letter-spacing: -.02em; }
.psperf { flex: 1.4; border-radius: 7px; padding: 7px 11px; display: flex; align-items: center; gap: 8px; }
.psperf .lp { font-size: 7px; font-weight: 700; letter-spacing: .07em; text-transform: uppercase; margin-bottom: 2px; }
.psperf .vp { font-size: 18px; font-weight: 900; line-height: 1; }
.psperf-bar { flex: 1; }
.stcard { background: #fff; border: 1px solid #e2e8f0; border-radius: 11px; overflow: visible; margin-bottom: 13px; }
.sth { padding: 7px 13px; background: #f8fafc; border-bottom: 1px solid #e2e8f0; display: flex; align-items: center; justify-content: space-between; }
.sth span { font-size: 8px; font-weight: 800; letter-spacing: .1em; text-transform: uppercase; color: #64748b; }
.scnt { font-size: 8px; background: #e2e8f0; color: #475569; border-radius: 999px; padding: 2px 8px; font-weight: 700; }
table { width: 100%; border-collapse: collapse; table-layout: fixed; }
thead th { background: #1e293b; color: #cbd5e1; font-size: 7.5px; font-weight: 700; padding: 7px 7px; text-align: left; text-transform: uppercase; letter-spacing: .05em; border-right: 1px solid #334155; }
thead th:last-child { border-right: none; }
tbody td { padding: 6px 7px; font-size: 8.5px; border-bottom: 1px solid #f1f5f9; border-right: 1px solid #f1f5f9; vertical-align: middle; color: #1e293b; word-break: break-word; }
tbody td:last-child { border-right: none; }
tbody tr:last-child td { border-bottom: none; }
tbody tr:nth-child(even) td { background: #f8fafc; }
.sno { font-weight: 800; color: #1e3a8a; }
.ptm { font-weight: 700; color: #0f172a; }
.pts { font-size: 7.5px; color: #94a3b8; margin-top: 2px; }
.badge { display: inline-flex; align-items: center; gap: 3px; padding: 2px 7px; border-radius: 999px; font-size: 7.5px; font-weight: 800; white-space: nowrap; }
.badge::before { content: ''; width: 5px; height: 5px; border-radius: 50%; display: inline-block; flex-shrink: 0; }
.bs { background: #dcfce7; color: #166534; }
.bs::before { background: #16a34a; }
.bw { background: #fff7ed; color: #9a3412; }
.bw::before { background: #ea580c; }
.bd { background: #fee2e2; color: #991b1b; }
.bd::before { background: #dc2626; }
.bn { background: #f1f5f9; color: #475569; }
.bn::before { background: #64748b; }
.fn { color: #dc2626; font-weight: 800; }
.fp { color: #15803d; font-weight: 800; }
.fz { color: #94a3b8; }
.footer { margin-top: 8px; display: flex; justify-content: space-between; font-size: 8px; color: #94a3b8; padding: 0 2px; }
.pdiv, .pozet, .sth { break-inside: avoid; page-break-inside: avoid; }
.stcard { break-inside: auto; page-break-inside: auto; }
thead { display: table-header-group; }
tfoot { display: table-footer-group; }
tr { break-inside: avoid; page-break-inside: avoid; }
tbody td { break-inside: avoid; page-break-inside: avoid; }
`;

const buildGlobalKpi = (summaries, bolge) => {
    const t = summaries.reduce(
        (a, s) => {
            a.talep += Number(s.talep || 0);
            a.tedarik += Number(s.tedarik || 0);
            a.edilmeyen += Number(s.edilmeyen || 0);
            a.gec += Number(s.gec_tedarik || 0);
            a.sho_b += Number(s.sho_basilan || 0);
            a.spot += Number(s.spot || 0);
            a.filo += Number(s.filo || 0);
            return a;
        },
        { talep: 0, tedarik: 0, edilmeyen: 0, gec: 0, sho_b: 0, spot: 0, filo: 0 }
    );

    const zam = Math.max(0, t.tedarik - t.gec);
    const perf =
        t.talep > 0
            ? Math.max(0, Math.min(100, Math.round((zam / t.talep) * 100)))
            : 0;
    const sho =
        t.tedarik > 0 ? Math.round((t.sho_b / t.tedarik) * 100) : 0;
    const ps = perfStyle(perf);

    return `
    <div class="gkpi">
      <div class="gkpi-h">&#127758; ${escapeHtml(bolge || "Genel")} &mdash; Genel Özet</div>
      <div class="krow">
        <div class="kp kp-b">
          <div class="l">Talep</div>
          <div class="v">${t.talep}</div>
        </div>
        <div class="kp kp-g">
          <div class="l">Tedarik</div>
          <div class="v">${t.tedarik}</div>
        </div>
        <div class="kp ${t.edilmeyen > 0 ? "kp-r" : "kp-n"}">
          <div class="l">Edilmeyen</div>
          <div class="v">${t.edilmeyen > 0 ? t.edilmeyen : "&mdash;"}</div>
        </div>
        <div class="kp ${t.gec > 0 ? "kp-a" : "kp-n"}">
          <div class="l">Ge&ccedil; Tedarik</div>
          <div class="v">${t.gec > 0 ? t.gec : "&mdash;"}</div>
        </div>
        <div class="kp kp-p">
          <div class="l">Spot</div>
          <div class="v">${t.spot}</div>
        </div>
        <div class="kp kp-n">
          <div class="l">Filo</div>
          <div class="v">${t.filo}</div>
        </div>
        <div class="pkp" style="background:${ps.bg};border:1px solid ${ps.border};">
          <div class="l" style="color:${ps.color};">Performans</div>
          <div class="v" style="color:${ps.color};">${perf}%</div>
          <div class="bw2"><div class="br" style="width:${perf}%;background:${ps.bar};"></div></div>
        </div>
        <div class="pkp" style="background:#f0f9ff;border:1px solid #bae6fd;">
          <div class="l" style="color:#0369a1;">SH&Ouml; Oran</div>
          <div class="v" style="color:#0369a1;">${sho}%</div>
          <div class="bw2"><div class="br" style="width:${sho}%;background:#0ea5e9;"></div></div>
        </div>
      </div>
    </div>
  `;
};

const buildProjeBlock = (summary, rows, dotColor) => {
    const ozet = hesaplaOzet(summary, rows);

    const p = Number(ozet.talep || 0);
    const t = Number(ozet.tedarik || 0);
    const ed = Number(ozet.edilmeyen || 0);
    const gec = Number(ozet.gec_tedarik || 0);
    const sho = Number(ozet.sho_basilan || 0);
    const zam = Math.max(0, t - gec);
    const oran = p > 0 ? Math.max(0, Math.min(100, Math.round((zam / p) * 100))) : 0;
    const ps = perfStyle(oran);

    const edHtml =
        ed > 0
            ? `<div class="ps kp-r"><div class="l">Edilmeyen</div><div class="v">${ed}</div></div>`
            : `<div class="ps kp-n"><div class="l">Edilmeyen</div><div class="v">&mdash;</div></div>`;

    const gecHtml =
        gec > 0
            ? `<div class="ps kp-a"><div class="l">Ge&ccedil; Tedarik</div><div class="v">${gec}</div></div>`
            : `<div class="ps kp-n"><div class="l">Ge&ccedil; Tedarik</div><div class="v">&mdash;</div></div>`;

    const siraliRows = [...rows].sort((a, b) => {
        const seferA = String(a?.seferNo || "").trim();
        const seferB = String(b?.seferNo || "").trim();
        const durumA = statusText(a);
        const durumB = statusText(b);

        const grup = (sefer, durum) => {
            const planlanmadi = sefer.toLocaleLowerCase("tr-TR") === "planlanmadı";
            if (planlanmadi && durum === "Yükleme Tarihi Yok") return 1;
            if (!planlanmadi && durum === "Zamanında") return 2;
            if (!planlanmadi && durum === "Geç Tedarik") return 3;
            return 4;
        };

        const gA = grup(seferA, durumA);
        const gB = grup(seferB, durumB);
        if (gA !== gB) return gA - gB;
        return seferA.localeCompare(seferB, "tr-TR");
    });

    const rowsHtml = siraliRows
        .map((r) => {
            const sc = statusClass(r);
            const st = statusText(r);
            const fark = r?.farkSaat;

            let farkHtml = `<span class="fz">&mdash;</span>`;
            if (fark !== null && fark !== undefined && fark !== "-") {
                const n = parseFloat(fark);
                if (!isNaN(n)) {
                    farkHtml =
                        n < 0
                            ? `<span class="fn">${n}</span>`
                            : n > 0
                                ? `<span class="fp">+${n}</span>`
                                : `<span class="fz">0</span>`;
                }
            }

            return `
        <tr>
          <td><span class="sno">${escapeHtml(r?.seferNo || "-")}</span></td>
          <td>${escapeHtml(r?.talepNo || "-")}</td>
          <td><div class="ptm">${escapeHtml(short(r?.musteri))}</div></td>
          <td>${escapeHtml(short(getYukleme(r)))}</td>
          <td>${escapeHtml(short(getTeslim(r)))}</td>
          <td>${escapeHtml(r?.yuklemeTarihi ? formatDateLong(r.yuklemeTarihi) : "-")}</td>
          <td>${escapeHtml(r?.yuklemeyeGelis ? formatDateLong(r.yuklemeyeGelis) : "-")}</td>
          <td>${farkHtml}</td>
          <td><span class="badge ${sc}">${escapeHtml(st)}</span></td>
          <td>${escapeHtml(r?.aracTipi || "-")}</td>
        </tr>
      `;
        })
        .join("\n");

    return `
    <div class="pdiv">
      <div class="pdiv-line"></div>
      <div class="pdiv-badge">
        <div class="pdiv-dot" style="background:${dotColor};"></div>
        ${escapeHtml(ozet.proje || "Proje")}
      </div>
      <div class="pdiv-line"></div>
    </div>

    <div class="pozet">
      <div class="ps kp-b"><div class="l">Talep</div><div class="v">${p}</div></div>
      <div class="ps kp-g"><div class="l">Tedarik</div><div class="v">${t}</div></div>
      ${edHtml}
      ${gecHtml}
      <div class="ps kp-p"><div class="l">Spot</div><div class="v">${ozet.spot}</div></div>
      <div class="ps kp-n"><div class="l">Filo</div><div class="v">${ozet.filo}</div></div>
      <div class="psperf" style="background:${ps.bg};border:1px solid ${ps.border};">
        <div>
          <div class="lp" style="color:${ps.color};">Performans</div>
          <div class="vp" style="color:${ps.color};">${oran}%</div>
        </div>
        <div class="psperf-bar">
          <div class="bw2"><div class="br" style="width:${oran}%;background:${ps.bar};"></div></div>
        </div>
      </div>
    </div>

    <div class="stcard">
      <div class="sth">
        <span>${escapeHtml(ozet.proje || "Proje")} &mdash; Sefer Listesi</span>
        <span class="scnt">${rows.length} sefer</span>
      </div>
      <table>
        <thead>
          <tr>
            <th style="width:9%;">Sefer No</th>
            <th style="width:8%;">Talep No</th>
            <th style="width:14%;">M&uuml;şteri</th>
            <th style="width:13%;">Y&uuml;kleme Noktası</th>
            <th style="width:13%;">Teslim Noktası</th>
            <th style="width:11%;">Y&uuml;kleme Tarihi</th>
            <th style="width:11%;">Y&uuml;klemeye Geli&#351;</th>
            <th style="width:7%;">Fark (s)</th>
            <th style="width:9%;">Durum</th>
            <th style="width:5%;">Ara&ccedil; Tipi</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml || `<tr><td colspan="10" style="text-align:center;color:#94a3b8;">Sefer bulunamadı.</td></tr>`}
        </tbody>
      </table>
    </div>
  `;
};

const buildHtml = (summaries, data, bolge) => {
    const guncelSummaries = summaries.map((s) => {
        const key = String(s.proje || "").trim().toLowerCase();
        const rows = data.filter(
            (r) => String(r.proje || "").trim().toLowerCase() === key
        );
        return hesaplaOzet(s, rows);
    });

    const bloklar = guncelSummaries
        .map((s, i) => {
            const key = String(s.proje || "").trim().toLowerCase();
            const rows = data.filter(
                (r) => String(r.proje || "").trim().toLowerCase() === key
            );
            return buildProjeBlock(s, rows, DOT_COLORS[i % DOT_COLORS.length]);
        })
        .join("\n");

    return `<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <title>${escapeHtml(bolge || "Sefer Analiz Raporu")}</title>
  <style>${CSS}</style>
</head>
<body>
  <div class="wrap">
    <div class="hero">
      <div>
        <div class="hb">Tedarik Analiz Raporu</div>
        <div class="ht">${escapeHtml(bolge || "Sefer Raporu")}</div>
        <div class="hs">Otomatik oluşturulmuş sefer analiz raporu</div>
      </div>
      <div class="hd">
        <div>${formatDateLong(new Date())}</div>
        <div>${guncelSummaries.length} proje &bull; ${data.length} sefer</div>
      </div>
    </div>
    ${buildGlobalKpi(guncelSummaries, bolge)}
    ${bloklar}
    <div class="footer">
      <span>Odak Lojistik &mdash; Tedarik Analiz Sistemi</span>
      <span>Olu&#351;turulma: ${formatDateLong(new Date())}</span>
    </div>
  </div>
</body>
</html>`;
};

/* ======================= EXCEL (MODERN) ======================= */

const buildExcelBuffer = ({ item, bolge }) => {
    const { data = [], summaries = [] } = item || {};

    // ─── STYLE HELPERS ─────────────────────────────────────────────
    const clr = {
        dark: "0F172A", dark2: "1E293B", slate: "334155", mid: "64748B",
        light: "94A3B8", white: "FFFFFF",
        blue: "1D4ED8", blueBg: "DBEAFE", blueBdr: "BFDBFE",
        green: "166534", greenBg: "DCFCE7", greenBdr: "BBF7D0",
        red: "991B1B", redBg: "FEE2E2", redBdr: "FECACA",
        amber: "92400E", amberBg: "FEF3C7",
        purple: "6B21A8", purpleBg: "FAF5FF",
        sky: "0369A1", skyBg: "F0F9FF",
        row0: "F8FAFC", row1: "FFFFFF",
        headerBg: "1E293B", sectionBg: "EFF6FF",
    };

    const mkFont = (bold, sz, rgb, mono) => ({
        name: mono ? "Courier New" : "Arial",
        bold: !!bold, sz: sz || 10,
        color: { rgb: rgb || clr.dark },
    });
    const mkFill = (rgb) => ({ type: "pattern", patternType: "solid", fgColor: { rgb } });
    const mkBorder = (style = "thin", rgb = "E2E8F0") => ({
        top: { style, color: { rgb } }, bottom: { style, color: { rgb } },
        left: { style, color: { rgb } }, right: { style, color: { rgb } },
    });
    const mkAlign = (h = "center", v = "center", wrap = false) => ({
        horizontal: h, vertical: v, wrapText: wrap,
    });

    const mergeAdd = (arr, r1, c1, r2, c2) =>
        arr.push({ s: { r: r1, c: c1 }, e: { r: r2, c: c2 } });

    const setCell = (ws, r, c, v, s) => {
        ws[XLSX.utils.encode_cell({ r, c })] = { v: v ?? "", t: "s", s };
    };
    const setNum = (ws, r, c, v, s) => {
        ws[XLSX.utils.encode_cell({ r, c })] = { v: Number(v) || 0, t: "n", s };
    };
    const setBlank = (ws, r, c, s) => {
        ws[XLSX.utils.encode_cell({ r, c })] = { v: "", t: "s", s };
    };

    const barChart = (pct, w = 20) => {
        const filled = Math.round((Math.min(100, Math.max(0, pct)) / 100) * w);
        return "█".repeat(filled) + "░".repeat(w - filled);
    };

    const perfColors = (p) =>
        p >= 90 ? { bg: "F0FDF4", fg: "166534" }
            : p >= 70 ? { bg: "DBEAFE", fg: "1D4ED8" }
                : p >= 50 ? { bg: "FEF3C7", fg: "92400E" }
                    : { bg: "FEE2E2", fg: "991B1B" };

    const fmtDate = (v) => {
        if (!v) return "-";
        const d = new Date(v);
        return isNaN(d.getTime()) ? String(v) : d.toLocaleDateString("tr-TR");
    };

    const isBosDeger = (v) => {
        const s = String(v ?? "")
            .trim()
            .toLocaleLowerCase("tr-TR");

        return (
            !s ||
            s === "-" ||
            s === "null" ||
            s === "undefined"
        );
    };

    const parseTarih = (v) => {
        if (!v) return null;

        if (v instanceof Date && !isNaN(v.getTime())) {
            return v;
        }

        const s = String(v).trim();

        const trMatch = s.match(
            /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/
        );

        if (trMatch) {
            const [, gun, ay, yil, saat = "0", dakika = "0"] = trMatch;

            return new Date(
                Number(yil),
                Number(ay) - 1,
                Number(gun),
                Number(saat),
                Number(dakika),
                0,
                0
            );
        }

        const d = new Date(s);

        return isNaN(d.getTime()) ? null : d;
    };

    const sonrakiIsGunuSaat6 = (tarih) => {
        if (!tarih) return null;

        const d = parseTarih(tarih);

        if (!d) return null;

        const next = new Date(d);

        next.setDate(next.getDate() + 1);

        while (next.getDay() === 0 || next.getDay() === 6) {
            next.setDate(next.getDate() + 1);
        }

        next.setHours(6, 0, 0, 0);

        return next;
    };

    const hesaplaDurumExcel = (row) => {
        const seferNo = String(row?.seferNo || "")
            .trim()
            .toLocaleLowerCase("tr-TR");

        if (
            seferNo.includes("planlamada") ||
            seferNo.includes("planlanmadı")
        ) {
            return "Tedarik Edilemeyen";
        }

        if (isBosDeger(row?.yuklemeyeGelis)) {
            return "Tedarik Edilemeyen";
        }

        const yuklemeTarihi = parseTarih(row?.yuklemeTarihi);

        const yuklemeyeGelis = parseTarih(row?.yuklemeyeGelis);

        if (!yuklemeTarihi || !yuklemeyeGelis) {
            return "Tedarik Edilemeyen";
        }

        const limit = sonrakiIsGunuSaat6(yuklemeTarihi);

        if (limit && yuklemeyeGelis > limit) {
            return "Geç Tedarik";
        }

        return "Zamanında";
    };

    // ─── AGGREGATES ────────────────────────────────────────────────
    const totalTalep = summaries.reduce((a, s) => a + Number(s.talep || 0), 0);
    const totalTedarik = summaries.reduce((a, s) => a + Number(s.tedarik || 0), 0);
    const totalEdilmeyen = summaries.reduce((a, s) => a + Number(s.edilmeyen || 0), 0);
    const totalGec = data.filter(
        (d) => hesaplaDurumExcel(d) === "Geç Tedarik"
    ).length;    const totalSpot = summaries.reduce((a, s) => a + Number(s.spot || 0), 0);
    const totalFilo = summaries.reduce((a, s) => a + Number(s.filo || 0), 0);
    const totalSho = summaries.reduce((a, s) => a + Number(s.sho_basilan || 0), 0);
    const totalZam = Math.max(0, totalTedarik - totalGec);
    const totalPerf = totalTalep > 0 ? Math.round((totalZam / totalTalep) * 100) : 0;
    const totalShoOran = totalTedarik > 0 ? Math.round((totalSho / totalTedarik) * 100) : 0;
    const durumlar = data.map((d) => hesaplaDurumExcel(d));

    const zamCount = durumlar.filter((d) => d === "Zamanında").length;

    const gecCount = durumlar.filter((d) => d === "Geç Tedarik").length;

    const yokCount = durumlar.filter(
        (d) => d === "Tedarik Edilemeyen"
    ).length;
    const wb = XLSX.utils.book_new();

    // ═══════════════════════════════════════════════════════════════
    // SHEET 1 — DASHBOARD
    // ═══════════════════════════════════════════════════════════════
    const ws1 = {}, m1 = [];
    let r = 0;

    // Banner
    mergeAdd(m1, r, 0, r, 11);
    setCell(ws1, r, 0,
        `  🚚 TEDARİK ANALİZ RAPORU  —  ${bolge || "TÜM BÖLGELER"}  |  Odak Lojistik`,
        { font: mkFont(true, 16, clr.white), fill: mkFill(clr.dark), alignment: mkAlign("left"), border: mkBorder("medium", clr.slate) }
    );
    for (let c = 1; c <= 11; c++) setBlank(ws1, r, c, { fill: mkFill(clr.dark) });
    r++;

    mergeAdd(m1, r, 0, r, 11);
    setCell(ws1, r, 0,
        `  Oluşturulma: ${new Date().toLocaleDateString("tr-TR", { day: "2-digit", month: "long", year: "numeric" })}   •   ${summaries.length} proje   •   ${data.length} sefer`,
        { font: mkFont(false, 9, clr.light), fill: mkFill(clr.dark2), alignment: mkAlign("left") }
    );
    for (let c = 1; c <= 11; c++) setBlank(ws1, r, c, { fill: mkFill(clr.dark2) });
    r++;
    r++;

    // KPI section header
    mergeAdd(m1, r, 0, r, 11);
    setCell(ws1, r, 0, "  📊  GENEL ÖZET — KPI GÖSTERGELERİ", {
        font: mkFont(true, 11, clr.dark), fill: mkFill(clr.sectionBg),
        alignment: mkAlign("left"),
        border: { bottom: { style: "medium", color: { rgb: clr.blue } } },
    });
    for (let c = 1; c <= 11; c++) setBlank(ws1, r, c, {
        fill: mkFill(clr.sectionBg),
        border: { bottom: { style: "medium", color: { rgb: clr.blue } } },
    });
    r++;
    r++;

    // KPI cards (label / big number / subtitle — 3 rows, 2 cols each)
    const kpiCards = [
        { label: "📦 TALEP", val: totalTalep, sub: "Toplam Talep", bg: clr.blueBg, fg: clr.blue, bdr: clr.blueBdr },
        { label: "✅ TEDARİK", val: totalTedarik, sub: "Temin Edilen", bg: clr.greenBg, fg: clr.green, bdr: clr.greenBdr },
        { label: "❌ EDİLMEYEN", val: totalEdilmeyen, sub: "Karşılanamayan", bg: totalEdilmeyen > 0 ? clr.redBg : "F1F5F9", fg: totalEdilmeyen > 0 ? clr.red : clr.mid, bdr: totalEdilmeyen > 0 ? clr.redBdr : "E2E8F0" },
        { label: "⏰ GEÇ TEDARİK", val: totalGec, sub: "Gecikmiş Sefer", bg: totalGec > 0 ? clr.amberBg : "F1F5F9", fg: totalGec > 0 ? clr.amber : clr.mid, bdr: totalGec > 0 ? "FED7AA" : "E2E8F0" },
        { label: "🚐 SPOT", val: totalSpot, sub: "Spot Araç", bg: clr.purpleBg, fg: clr.purple, bdr: "E9D5FF" },
        { label: "🚛 FİLO", val: totalFilo, sub: "Filo Aracı", bg: "F0FDF4", fg: "14532D", bdr: "A7F3D0" },
    ];
    const kpiCols = [0, 2, 4, 6, 8, 10];

    kpiCards.forEach((k, i) => {
        const col = kpiCols[i];
        const bT = { style: "medium", color: { rgb: k.bdr } };
        const bB = { style: "medium", color: { rgb: k.bdr } };
        const bL = { style: "medium", color: { rgb: k.bdr } };
        const bR = { style: "medium", color: { rgb: k.bdr } };

        mergeAdd(m1, r, col, r, col + 1);
        setCell(ws1, r, col, k.label, { font: mkFont(true, 8, k.fg), fill: mkFill(k.bg), alignment: mkAlign("center"), border: { top: bT, left: bL, right: bR } });
        setBlank(ws1, r, col + 1, { fill: mkFill(k.bg), border: { top: bT, right: bR } });

        mergeAdd(m1, r + 1, col, r + 1, col + 1);
        setNum(ws1, r + 1, col, k.val, { font: mkFont(true, 24, k.fg), fill: mkFill(k.bg), alignment: mkAlign("center"), border: { left: bL, right: bR } });
        setBlank(ws1, r + 1, col + 1, { fill: mkFill(k.bg), border: { right: bR } });

        mergeAdd(m1, r + 2, col, r + 2, col + 1);
        setCell(ws1, r + 2, col, k.sub, { font: mkFont(false, 8, k.fg), fill: mkFill(k.bg), alignment: mkAlign("center"), border: { bottom: bB, left: bL, right: bR } });
        setBlank(ws1, r + 2, col + 1, { fill: mkFill(k.bg), border: { bottom: bB, right: bR } });
    });
    r += 3;
    r++;

    // Performans + SHÖ bar kartları
    const perfC = perfColors(totalPerf);
    const perfCards2 = [
        {
            label: "⚡ PERFORMANS ORANI",
            val: totalPerf,
            bar: barChart(totalPerf),
            sub: `${totalZam} zamanında / ${totalTalep} toplam talep`,
            bg: perfC.bg, fg: perfC.fg, bdr: perfC.fg,
        },
        {
            label: "📋 SHÖ BASIM ORANI",
            val: totalShoOran,
            bar: barChart(totalShoOran),
            sub: `${totalSho} SHÖ basılan / ${totalTedarik} tedarik`,
            bg: clr.skyBg, fg: clr.sky, bdr: "7DD3FC",
        },
    ];

    perfCards2.forEach((pc, i) => {
        const cs = i * 6;
        const bdr = { style: "medium", color: { rgb: pc.bdr } };

        mergeAdd(m1, r, cs, r, cs + 5);
        setCell(ws1, r, cs, pc.label, { font: mkFont(true, 9, pc.fg), fill: mkFill(pc.bg), alignment: mkAlign("left"), border: { top: bdr, left: bdr, right: bdr } });
        for (let c = cs + 1; c <= cs + 5; c++) setBlank(ws1, r, c, { fill: mkFill(pc.bg), border: { top: bdr, right: bdr } });

        mergeAdd(m1, r + 1, cs, r + 1, cs + 1);
        setCell(ws1, r + 1, cs, `${pc.val}%`, { font: mkFont(true, 28, pc.fg), fill: mkFill(pc.bg), alignment: mkAlign("center"), border: { left: bdr } });
        setBlank(ws1, r + 1, cs + 1, { fill: mkFill(pc.bg) });
        mergeAdd(m1, r + 1, cs + 2, r + 1, cs + 5);
        setCell(ws1, r + 1, cs + 2, pc.bar, { font: mkFont(false, 11, pc.fg, true), fill: mkFill(pc.bg), alignment: mkAlign("left", "center"), border: { right: bdr } });
        for (let c = cs + 3; c <= cs + 5; c++) setBlank(ws1, r + 1, c, { fill: mkFill(pc.bg), border: { right: bdr } });

        mergeAdd(m1, r + 2, cs, r + 2, cs + 5);
        setCell(ws1, r + 2, cs, pc.sub, { font: mkFont(false, 8, pc.fg), fill: mkFill(pc.bg), alignment: mkAlign("left"), border: { bottom: bdr, left: bdr, right: bdr } });
        for (let c = cs + 1; c <= cs + 5; c++) setBlank(ws1, r + 2, c, { fill: mkFill(pc.bg), border: { bottom: bdr, right: bdr } });
    });
    r += 3;
    r++;

    // Proje tablosu başlığı
    mergeAdd(m1, r, 0, r, 11);
    setCell(ws1, r, 0, "  📁  PROJE BAZLI PERFORMANS TABLOSU", {
        font: mkFont(true, 11, clr.dark), fill: mkFill(clr.sectionBg),
        alignment: mkAlign("left"),
        border: { bottom: { style: "medium", color: { rgb: clr.blue } } },
    });
    for (let c = 1; c <= 11; c++) setBlank(ws1, r, c, {
        fill: mkFill(clr.sectionBg),
        border: { bottom: { style: "medium", color: { rgb: clr.blue } } },
    });
    r++;
    r++;

    ["Proje", "Talep", "Tedarik", "Edilmeyen", "Geç Tedarik", "Zamanında", "Perf %", "Grafik", "Spot", "Filo", "SHÖ Basılan", "SHÖ Oranı"].forEach((h, c) => {
        setCell(ws1, r, c, h, {
            font: mkFont(true, 9, clr.white), fill: mkFill(clr.headerBg),
            alignment: mkAlign("center"), border: mkBorder("thin", clr.slate),
        });
    });
    r++;

    summaries.forEach((s, idx) => {
        const tal = Number(s.talep || 0);
        const ted = Number(s.tedarik || 0);
        const ed = Number(s.edilmeyen || 0);
        const projeRows = data.filter(
            (d) =>
                String(d.proje || "").trim().toLocaleLowerCase("tr-TR") ===
                String(s.proje || "").trim().toLocaleLowerCase("tr-TR")
        );

        const gec2 = projeRows.filter(
            (d) => hesaplaDurumExcel(d) === "Geç Tedarik"
        ).length;
        const zam2 = Math.max(0, ted - gec2);
        const perf2 = tal > 0 ? Math.round((zam2 / tal) * 100) : 0;
        const sho2 = Number(s.sho_basilan || 0);
        const shoO = ted > 0 ? Math.round((sho2 / ted) * 100) : 0;
        const pc2 = perfColors(perf2);
        const rb = idx % 2 === 0 ? clr.row0 : clr.row1;
        const rs = (bg) => ({ font: mkFont(false, 10, clr.dark), fill: mkFill(bg || rb), alignment: mkAlign("center"), border: mkBorder("thin", "E2E8F0") });

        setCell(ws1, r, 0, s.proje || "-", { ...rs(), font: mkFont(true, 10, clr.dark2), alignment: mkAlign("left") });
        setNum(ws1, r, 1, tal, { ...rs(clr.blueBg), font: mkFont(true, 10, clr.blue) });
        setNum(ws1, r, 2, ted, { ...rs(clr.greenBg), font: mkFont(true, 10, clr.green) });
        setNum(ws1, r, 3, ed, { ...rs(ed > 0 ? clr.redBg : rb), font: mkFont(ed > 0, 10, ed > 0 ? clr.red : clr.mid) });
        setNum(ws1, r, 4, gec2, { ...rs(gec2 > 0 ? clr.amberBg : rb), font: mkFont(gec2 > 0, 10, gec2 > 0 ? clr.amber : clr.mid) });
        setNum(ws1, r, 5, zam2, rs());
        setCell(ws1, r, 6, `${perf2}%`, { font: mkFont(true, 10, pc2.fg), fill: mkFill(pc2.bg), alignment: mkAlign("center"), border: mkBorder("thin", "E2E8F0") });
        setCell(ws1, r, 7, barChart(perf2, 12), { font: mkFont(false, 10, pc2.fg, true), fill: mkFill(pc2.bg), alignment: mkAlign("left", "center"), border: mkBorder("thin", "E2E8F0") });
        setNum(ws1, r, 8, Number(s.spot || 0), rs());
        setNum(ws1, r, 9, Number(s.filo || 0), rs());
        setNum(ws1, r, 10, sho2, rs());
        setCell(ws1, r, 11, `${shoO}%`, { font: mkFont(false, 10, clr.sky), fill: mkFill(clr.skyBg), alignment: mkAlign("center"), border: mkBorder("thin", "E2E8F0") });
        r++;
    });

    // Toplam satırı
    const ts = { font: mkFont(true, 10, clr.white), fill: mkFill(clr.dark2), alignment: mkAlign("center"), border: mkBorder("medium", clr.slate) };
    setCell(ws1, r, 0, "TOPLAM", { ...ts, alignment: mkAlign("left") });
    setNum(ws1, r, 1, totalTalep, ts);
    setNum(ws1, r, 2, totalTedarik, ts);
    setNum(ws1, r, 3, totalEdilmeyen, ts);
    setNum(ws1, r, 4, totalGec, ts);
    setNum(ws1, r, 5, totalZam, ts);
    setCell(ws1, r, 6, `${totalPerf}%`, { ...ts, font: mkFont(true, 10, totalPerf >= 70 ? "86EFAC" : "FCA5A5") });
    setCell(ws1, r, 7, barChart(totalPerf, 12), {
        ...ts,
        font: { name: "Courier New", sz: 10, color: { rgb: totalPerf >= 70 ? "86EFAC" : "FCA5A5" }, bold: true },
    });
    setNum(ws1, r, 8, totalSpot, ts);
    setNum(ws1, r, 9, totalFilo, ts);
    setNum(ws1, r, 10, totalSho, ts);
    setCell(ws1, r, 11, `${totalShoOran}%`, { ...ts, font: mkFont(true, 10, "7DD3FC") });
    r++;
    r++;

    // Durum dağılımı
    mergeAdd(m1, r, 0, r, 11);
    setCell(ws1, r, 0, "  📊  DURUM DAĞILIMI", {
        font: mkFont(true, 11, clr.dark), fill: mkFill(clr.sectionBg),
        alignment: mkAlign("left"),
        border: { bottom: { style: "medium", color: { rgb: clr.blue } } },
    });
    for (let c = 1; c <= 11; c++) setBlank(ws1, r, c, {
        fill: mkFill(clr.sectionBg),
        border: { bottom: { style: "medium", color: { rgb: clr.blue } } },
    });
    r++;
    r++;

    [
        { label: "✅ Zamanında", count: zamCount, pct: data.length ? Math.round((zamCount / data.length) * 100) : 0, bg: clr.greenBg, fg: clr.green },
        { label: "⏰ Geç Tedarik", count: gecCount, pct: data.length ? Math.round((gecCount / data.length) * 100) : 0, bg: clr.amberBg, fg: clr.amber },
        { label: "❓ Yükleme Tarihi Yok", count: yokCount, pct: data.length ? Math.round((yokCount / data.length) * 100) : 0, bg: "FFF7ED", fg: "9A3412" },
    ].forEach((si) => {
        setCell(ws1, r, 0, si.label, { font: mkFont(true, 10, si.fg), fill: mkFill(si.bg), alignment: mkAlign("left"), border: mkBorder("thin", "E2E8F0") });
        setNum(ws1, r, 1, si.count, { font: mkFont(true, 12, si.fg), fill: mkFill(si.bg), alignment: mkAlign("center"), border: mkBorder("thin", "E2E8F0") });
        setCell(ws1, r, 2, `${si.pct}%`, { font: mkFont(true, 10, si.fg), fill: mkFill(si.bg), alignment: mkAlign("center"), border: mkBorder("thin", "E2E8F0") });
        mergeAdd(m1, r, 3, r, 11);
        setCell(ws1, r, 3, barChart(si.pct, 30), { font: mkFont(false, 10, si.fg, true), fill: mkFill(si.bg), alignment: mkAlign("left", "center"), border: mkBorder("thin", "E2E8F0") });
        for (let c = 4; c <= 11; c++) setBlank(ws1, r, c, { fill: mkFill(si.bg), border: mkBorder("thin", "E2E8F0") });
        r++;
    });

    ws1["!cols"] = [
        { wch: 22 }, { wch: 9 }, { wch: 10 }, { wch: 12 }, { wch: 13 },
        { wch: 11 }, { wch: 13 }, { wch: 18 }, { wch: 9 }, { wch: 9 }, { wch: 13 }, { wch: 11 },
    ];
    ws1["!rows"] = [{ hpt: 30 }, { hpt: 18 }];
    ws1["!merges"] = m1;
    ws1["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r + 5, c: 11 } });
    ws1["!freeze"] = { xSplit: 0, ySplit: 3 };
    XLSX.utils.book_append_sheet(wb, ws1, "📊 Dashboard");

    // ═══════════════════════════════════════════════════════════════
    // SHEET 2 — ÖZET
    // ═══════════════════════════════════════════════════════════════
    const ws2 = {}, m2 = [];
    let r2 = 0;

    mergeAdd(m2, r2, 0, r2, 9);
    setCell(ws2, r2, 0, "  📁  PROJE ÖZET RAPORU", {
        font: mkFont(true, 13, clr.white), fill: mkFill(clr.dark), alignment: mkAlign("left"),
    });
    for (let c = 1; c <= 9; c++) setBlank(ws2, r2, c, { fill: mkFill(clr.dark) });
    r2++;
    r2++;

    ["Bölge", "Proje", "Talep", "Tedarik", "Tedarik Edilmeyen", "Filo", "Spot", "SHÖ Basılan", "Geç Tedarik", "Performans"].forEach((h, c) => {
        setCell(ws2, r2, c, h, {
            font: mkFont(true, 9, clr.white), fill: mkFill(clr.headerBg),
            alignment: mkAlign("center"), border: mkBorder("thin", clr.slate),
        });
    });
    r2++;

    summaries.forEach((s, idx) => {
        const tal = Number(s.talep || 0);
        const ted = Number(s.tedarik || 0);
        const gec2 = Number(s.gec_tedarik || 0);
        const zam2 = Math.max(0, ted - gec2);
        const perf2 = tal > 0 ? Math.round((zam2 / tal) * 100) : 0;
        const pc2 = perfColors(perf2);
        const rb = idx % 2 === 0 ? clr.row0 : clr.row1;
        const rs = (bg) => ({ font: mkFont(false, 10, clr.dark), fill: mkFill(bg || rb), alignment: mkAlign("center"), border: mkBorder("thin", "E2E8F0") });

        setCell(ws2, r2, 0, bolge || "-", rs());
        setCell(ws2, r2, 1, s.proje || "-", { ...rs(), font: mkFont(true, 10, clr.dark2), alignment: mkAlign("left") });
        setNum(ws2, r2, 2, tal, { ...rs(clr.blueBg), font: mkFont(true, 10, clr.blue) });
        setNum(ws2, r2, 3, ted, { ...rs(clr.greenBg), font: mkFont(true, 10, clr.green) });
        setNum(ws2, r2, 4, Number(s.edilmeyen || 0), rs(Number(s.edilmeyen || 0) > 0 ? clr.redBg : rb));
        setNum(ws2, r2, 5, Number(s.filo || 0), rs());
        setNum(ws2, r2, 6, Number(s.spot || 0), rs());
        setNum(ws2, r2, 7, Number(s.sho_basilan || 0), rs());
        setNum(ws2, r2, 8, gec2, { ...rs(gec2 > 0 ? clr.amberBg : rb), font: mkFont(gec2 > 0, 10, gec2 > 0 ? clr.amber : clr.dark) });
        setCell(ws2, r2, 9, `${perf2}% ${barChart(perf2, 8)}`, {
            font: mkFont(true, 10, pc2.fg, true), fill: mkFill(pc2.bg),
            alignment: mkAlign("left", "center"), border: mkBorder("thin", "E2E8F0"),
        });
        r2++;
    });

    ws2["!cols"] = [
        { wch: 14 }, { wch: 28 }, { wch: 8 }, { wch: 10 }, { wch: 18 },
        { wch: 8 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 20 },
    ];
    ws2["!merges"] = m2;
    ws2["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r2 + 2, c: 9 } });
    ws2["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 2, c: 0 }, e: { r: 2, c: 9 } }) };
    ws2["!freeze"] = { xSplit: 0, ySplit: 3 };
    ws2["!rows"] = [{ hpt: 26 }];
    XLSX.utils.book_append_sheet(wb, ws2, "📁 Özet");

    // ═══════════════════════════════════════════════════════════════
    // SHEET 3 — DETAY SEFERLER
    // ═══════════════════════════════════════════════════════════════
    const ws3 = {}, m3 = [];
    let r3 = 0;

    mergeAdd(m3, r3, 0, r3, 11);
    setCell(ws3, r3, 0, "  🗂️  DETAY SEFER LİSTESİ", {
        font: mkFont(true, 13, clr.white), fill: mkFill(clr.dark), alignment: mkAlign("left"),
    });
    for (let c = 1; c <= 11; c++) setBlank(ws3, r3, c, { fill: mkFill(clr.dark) });
    r3++;

    mergeAdd(m3, r3, 0, r3, 11);
    setCell(ws3, r3, 0,
        `  Toplam ${data.length} sefer  •  Zamanında: ${zamCount}  •  Geç: ${gecCount}  •  Tarihi Yok: ${yokCount}`,
        { font: mkFont(false, 9, clr.light), fill: mkFill(clr.dark2), alignment: mkAlign("left") }
    );
    for (let c = 1; c <= 11; c++) setBlank(ws3, r3, c, { fill: mkFill(clr.dark2) });
    r3++;
    r3++;

    ["Bölge", "Proje", "Sefer No", "Talep No", "Müşteri", "Yükleme Noktası", "Teslim Noktası", "Yükleme Tarihi", "Yüklemeye Geliş", "Fark (s)", "Durum", "Araç Tipi"].forEach((h, c) => {
        setCell(ws3, r3, c, h, {
            font: mkFont(true, 9, clr.white), fill: mkFill(clr.headerBg),
            alignment: mkAlign("center"), border: mkBorder("thin", clr.slate),
        });
    });
    r3++;



    data.forEach((row, idx) => {
        const rb = idx % 2 === 0 ? clr.row0 : clr.row1;
        const rs = (bg) => ({ font: mkFont(false, 9, clr.dark), fill: mkFill(bg || rb), alignment: mkAlign("center", "center", true), border: mkBorder("thin", "E2E8F0") });
        const durum = hesaplaDurumExcel(row);

        let durumBg = rb, durumFg = clr.dark;
        if (durum === "Zamanında") {
            durumBg = clr.greenBg;
            durumFg = clr.green;
        }
        else if (durum === "Geç Tedarik") {
            durumBg = clr.redBg;
            durumFg = clr.red;
        }
        else if (durum === "Tedarik Edilemeyen") {
            durumBg = clr.amberBg;
            durumFg = clr.amber;
        }
        const fark = row.farkSaat;
        const farkStr = (fark === null || fark === undefined || fark === "-") ? "-"
            : fark < 0 ? `${fark}` : fark > 0 ? `+${fark}` : "0";
        const farkFg = fark < 0 ? "15803D" : fark > 0 ? clr.red : clr.mid;


        setCell(ws3, r3, 0, row.bolge || bolge || "-", rs());
        setCell(ws3, r3, 1, row.proje || "-", { ...rs(), font: mkFont(true, 9, clr.dark2), alignment: mkAlign("left", "center", true) });
        setCell(ws3, r3, 2, row.seferNo || "-", { ...rs(), font: mkFont(true, 9, clr.blue) });
        setCell(ws3, r3, 3, row.talepNo || "-", rs());
        setCell(ws3, r3, 4, row.musteri || "-", { ...rs(), alignment: mkAlign("left", "center", true) });
        setCell(
            ws3,
            r3,
            5,
            formatNokta(
                row.yuklemeNoktasi,
                row.yuklemeIl,
                row.yuklemeIlce,
                row.yuklemeAdres,
                getYukleme(row)
            ),
            {
                ...rs(),
                alignment: mkAlign("left", "center", true)
            }
        );

        setCell(
            ws3,
            r3,
            6,
            formatNokta(
                row.teslimNoktasi,
                row.teslimIl,
                row.teslimIlce,
                row.teslimAdres,
                getTeslim(row)
            ),
            {
                ...rs(),
                alignment: mkAlign("left", "center", true)
            }
        );
        setCell(ws3, r3, 7, fmtDate(row.yuklemeTarihi), rs());
        setCell(ws3, r3, 8, fmtDate(row.yuklemeyeGelis), rs());
        setCell(ws3, r3, 9, farkStr, { font: mkFont(true, 10, farkStr === "-" ? clr.mid : farkFg), fill: mkFill(rb), alignment: mkAlign("center"), border: mkBorder("thin", "E2E8F0") });
        setCell(ws3, r3, 10, durum, {
            font: mkFont(true, 9, durumFg),
            fill: mkFill(durumBg),
            alignment: mkAlign("center"),
            border: mkBorder("thin", "E2E8F0")
        });

        setCell(ws3, r3, 11, row.aracTipi || "-", rs());
        r3++;
    });

    ws3["!cols"] = [
        { wch: 12 }, { wch: 22 }, { wch: 14 }, { wch: 12 }, { wch: 30 },
        { wch: 24 }, { wch: 24 }, { wch: 14 }, { wch: 14 }, { wch: 9 },
        { wch: 20 }, { wch: 12 },
    ];
    ws3["!merges"] = m3;
    ws3["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r3 + 2, c: 11 } });
    ws3["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 3, c: 0 }, e: { r: 3, c: 11 } }) };
    ws3["!freeze"] = { xSplit: 0, ySplit: 4 };
    ws3["!rows"] = [{ hpt: 26 }, { hpt: 18 }];
    XLSX.utils.book_append_sheet(wb, ws3, "🗂️ Detay Seferler");

    return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
};

/* ======================= MAIL GÖNDER ======================= */

const analizMailGonder = async ({ mailPayload, bolge }) => {
    let browser;

    try {
        if (!Array.isArray(mailPayload) || mailPayload.length === 0) {
            throw new Error("Gönderilecek mail verisi bulunamadı.");
        }

        const results = [];

        for (const item of mailPayload) {
            const { email, ccEmails = [], data = [], summaries = [] } = item || {};

            if (!email) {
                results.push({ email: "-", ok: false, message: "Email bulunamadı." });
                continue;
            }

            const effectiveSummaries =
                summaries.length > 0
                    ? summaries
                    : [
                        {
                            proje: bolge || "Sefer Raporu",
                            talep: data.length,
                            tedarik: data.length,
                            edilmeyen: 0,
                            gec_tedarik: 0,
                            sho_basilan: 0,
                            sho_basilmayan: 0,
                            spot: 0,
                            filo: 0,
                        },
                    ];

            const html = buildHtml(effectiveSummaries, data, bolge);

            browser = await puppeteer.launch({
                headless: true,
                args: [
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                ],
            });

            const page = await browser.newPage();
            await page.setContent(html, { waitUntil: "domcontentloaded", timeout: 120000 });
            await page.emulateMediaType("screen");
            await new Promise((resolve) => setTimeout(resolve, 500));

            const pdfBuffer = await page.pdf({
                format: "A4",
                landscape: true,
                printBackground: true,
                margin: { top: "8mm", right: "6mm", bottom: "8mm", left: "6mm" },
            });

            await browser.close();
            browser = null;

            console.log("📨 Mail gönderiliyor:", email);

            const info = await sendMailWithResend({
                to: email,
                cc: Array.isArray(ccEmails) && ccEmails.length > 0 ? ccEmails : [],
                subject: `📊 ${bolge || "Bölge"} | Sefer Analiz Raporu`,
                text: `Değerli Kullanıcılar,\n\nİlgili tarih aralığında oluşturmuş olduğunuz seferlere ait kontrollerinizi yapmanızı rica ederiz.\n\nSeferlerinizde eksik, fazla veya hatalı bir durum tespit etmeniz halinde müşteri hizmetleri birimimiz ile iletişime geçebilirsiniz.\n\nBilginize sunar, iyi günler dileriz.`,
                attachments: [
                    {
                        filename: `sefer-analiz-raporu-${slugify(bolge || "bolge")}.pdf`,
                        content: pdfBuffer,
                    },
                ],
            });

            results.push({ email, ok: true, messageId: info.messageId });
        }

        return results;
    } catch (err) {
        if (browser) await browser.close().catch(() => { });
        throw err;
    }
};

/* ======================= TMS ORDERS ======================= */

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const cache = new Map();
const CACHE_TTL_MS = 5 * 60 * 1000;

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

function extractFirstItem(data) {
    return (
        data?.data?.[0] ||
        data?.Data?.[0] ||
        data?.items?.[0] ||
        data?.[0] ||
        null
    );
}

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

        const key = cacheKey({ startDate, endDate, userId });
        const hit = cache.get(key);

        if (hit && Date.now() - hit.ts < CACHE_TTL_MS) {
            return res.json({ rid, ok: true, cached: true, data: hit.value });
        }

        const upstreamUrl = "https://api.odaklojistik.com.tr/api/tmsorders/getall";
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
            const firstItem = extractFirstItem(data);
            console.log("🔍 UPSTREAM FIRST ITEM:", firstItem);
            console.log("🔍 LOADING ALANLARI:", {
                TMSLoadingDocumentPrintedDate: firstItem?.TMSLoadingDocumentPrintedDate,
                TMSLoadingDocumentPrintedBy: firstItem?.TMSLoadingDocumentPrintedBy,
            });
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

/* ======================= ROUTES ======================= */

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
            "POST /send-analiz-mail",
            "POST /download-analiz-pdf",
            "POST /download-analiz-excel",
        ],
        allowedOrigins: ALLOWED_ORIGINS,
    });
});

app.post("/tmsorders", tmsordersHandler);
app.post("/tmsorders/week", tmsordersHandler);

app.post("/download-analiz-pdf", async (req, res) => {
    let browser;
    try {
        const { item, bolge } = req.body;

        if (!item) {
            return res.status(400).json({ ok: false, message: "PDF için item verisi bulunamadı." });
        }

        const { data = [], summaries = [] } = item;

        const effectiveSummaries =
            summaries.length > 0
                ? summaries
                : [
                    {
                        proje: bolge || "Sefer Raporu",
                        talep: data.length,
                        tedarik: data.length,
                        edilmeyen: 0,
                        gec_tedarik: 0,
                        sho_basilan: 0,
                        sho_basilmayan: 0,
                        spot: 0,
                        filo: 0,
                    },
                ];

        const html = buildHtml(effectiveSummaries, data, bolge);

        browser = await puppeteer.launch({
            headless: true,
            args: [
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
            ],
        });

        const page = await browser.newPage();
        await page.setContent(html, { waitUntil: "domcontentloaded", timeout: 120000 });
        await page.emulateMediaType("screen");
        await new Promise((resolve) => setTimeout(resolve, 500));

        const pdfBuffer = await page.pdf({
            format: "A4",
            landscape: true,
            printBackground: true,
            margin: { top: "8mm", right: "6mm", bottom: "8mm", left: "6mm" },
        });

        await browser.close();
        browser = null;

        const fileName = `sefer-analiz-raporu-${slugify(bolge || "bolge")}.pdf`;
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        return res.send(pdfBuffer);
    } catch (err) {
        if (browser) await browser.close().catch(() => { });
        console.error("PDF oluşturma hatası:", err);
        return res.status(500).json({ ok: false, message: err.message || "PDF oluşturulamadı." });
    }
});

app.post("/download-analiz-excel", async (req, res) => {
    try {
        const { item, bolge } = req.body;

        if (!item) {
            return res.status(400).json({ ok: false, message: "Excel için item verisi bulunamadı." });
        }

        const excelBuffer = buildExcelBuffer({ item, bolge });
        const fileName = `sefer-analiz-raporu-${slugify(bolge || "bolge")}.xlsx`;

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        return res.send(excelBuffer);
    } catch (err) {
        console.error("Excel oluşturma hatası:", err);
        return res.status(500).json({ ok: false, message: err.message || "Excel oluşturulamadı." });
    }
});

app.post("/send-analiz-mail", async (req, res) => {
    console.log("📩 GELEN BODY:", req.body);
    try {
        const { mailPayload, bolge } = req.body;

        sonPayloadKaydet({
            mailPayload,
            bolge,
            savedAt: new Date().toISOString(),
        });

        const results = await analizMailGonder({ mailPayload, bolge });

        return res.status(200).json({
            ok: true,
            message: "PDF rapor mailleri başarıyla gönderildi.",
            results,
        });
    } catch (err) {
        console.error("MAIL HATASI:", err);
        return res.status(500).json({
            ok: false,
            message: err.message || "Mail gönderilemedi.",
        });
    }
});

/* ======================= CRON ======================= */

const otomatikAnalizMailGonder = async () => {
    try {
        console.log("🟡 Otomatik analiz mail görevi başladı:", new Date());

        const saved = sonPayloadOku();

        if (!saved?.mailPayload?.length) {
            console.log("⚠️ Kayıtlı mailPayload yok. Önce ekrandan bir kez mail gönder.");
            return;
        }

        const results = await analizMailGonder({
            mailPayload: saved.mailPayload,
            bolge: saved.bolge || "TUM_BOLGELER",
        });

        console.log("✅ Otomatik mail gönderildi:", results);
    } catch (err) {
        console.error("❌ Otomatik analiz mail hatası:", err);
    }
};

cron.schedule(
    "0 10,17 * * *",
    async () => {
        await otomatikAnalizMailGonder();
    },
    { timezone: "Europe/Istanbul" }
);

/* ======================= START ======================= */

app.listen(PORT, () => {
    console.log("🚀 Server listening on port", PORT);
});