// index.js

const express = require("express"); const cors = require("cors"); const dotenv = require("dotenv"); const puppeteer = require("puppeteer"); const cron = require("node-cron"); const fs = require("fs"); const path = require("path"); const dns = require("dns"); const XLSX = require("xlsx-js-style");

dns.setDefaultResultOrder("ipv4first");

dotenv.config();

const app = express(); const PORT = process.env.PORT || 10000;

const ALLOWED_ORIGINS = ["http://localhost:3000", "http://127.0.0.1:3000", "http://localhost:3001", "http://127.0.0.1:3001", "http://localhost:5173", "http://127.0.0.1:5173", "https://analiz-pearl.vercel.app", "https://analiz-v2.vercel.app",];

const corsOptions = {
    origin: (origin, cb) => {
        if (!origin) return cb(null, true); if (ALLOWED_ORIGINS.includes(origin)) return cb(null, true);

        console.warn("⛔ CORS blocked origin:", origin);
        return cb(null, false);
    },
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: false,
    maxAge: 86400,

};

app.use(cors(corsOptions)); app.options(/.*/, cors(corsOptions)); app.use(express.json({ limit: "10mb" }));

/* =======================MAIL AYARLARI======================= */

const sendMailWithResend = async ({ to, cc, subject, text, attachments }) => {
    const body = { from: "Odak Lojistik onboarding@resend.dev", to: [to], subject, text, };

    if (cc && cc.length > 0) body.cc = cc;

    if (attachments && attachments.length > 0) {
        body.attachments = attachments.map(a => ({
            filename: a.filename,
            content: Buffer.from(a.content).toString("base64"),
        }));
    }

    const res = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${process.env.RESEND_API_KEY}`,
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

const sonPayloadKaydet = (payload) => { fs.writeFileSync(LAST_PAYLOAD_FILE, JSON.stringify(payload, null, 2), "utf8"); };

const sonPayloadOku = () => { if (!fs.existsSync(LAST_PAYLOAD_FILE)) return null; return JSON.parse(fs.readFileSync(LAST_PAYLOAD_FILE, "utf8")); };

const escapeHtml = (v) => String(v ?? "-").replace(/&/g, "&").replace(/</g, "<").replace(/>/g, ">").replace(/"/g, """).replace(/'/g, "'");

const slugify = (t) => String(t || "rapor").toLowerCase().trim().replace(/ğ/g, "g").replace(/ü/g, "u").replace(/ş/g, "s").replace(/ı/g, "i").replace(/ö/g, "o").replace(/ç/g, "c").replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");

const pick = (...vals) => vals.find((v) => v !== undefined && v !== null && String(v).trim() !== "");

const getYukleme = (r = {}) => pick(r.PickupAddressCode, r.yuklemeNoktasi, r.yuklemeNokta, r.yuklemeNoktaAdi, r.yuklemeYeri, r.yuklemeAdres, r.yuklemeLokasyon, r.cikisNoktasi, r.cikisYeri, r.gondericiUnvan, r.gonderici, r.gonderen, "-");

const getTeslim = (r = {}) => pick(r.DeliveryAddressCode, r.teslimNoktasi, r.teslimNokta, r.teslimNoktaAdi, r.teslimYeri, r.teslimAdres, r.teslimLokasyon, r.varisNoktasi, r.varisYeri, r.aliciUnvan, r.alici, "-");

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

const short = (t, max = 42) => { const s = String(t || "-"); return s.length > max ? s.slice(0, max) + "..." : s; };

const statusText = (r) => {
    if (!r?.yuklemeyeGelis || String(r.yuklemeyeGelis).trim() === "-") { return "Yükleme Tarihi Yok"; }

    return r?.durumText || "-";

};

const statusClass = (r) => {
    const s = statusText(r);

    if (s === "Zamanında") return "bs";
    if (s === "Geç Tedarik") return "bd";
    if (s === "Yükleme Tarihi Yok") return "bw";

    return "bn";

};

const hesaplaOzet = (summary = {}, rows = []) => { return { ...summary, talep: Number(summary.talep || 0), tedarik: Number(summary.tedarik || 0), edilmeyen: Number(summary.edilmeyen || 0), gec_tedarik: Number(summary.gec_tedarik || 0), sho_basilan: Number(summary.sho_basilan || 0), sho_basilmayan: Number(summary.sho_basilmayan || 0), spot: Number(summary.spot || 0), filo: Number(summary.filo || 0), }; };

const perfStyle = (p) => {
    if (p >= 90) return { bg: "#f0fdf4", border: "#bbf7d0", color: "#166534", bar: "#16a34a" }; if (p >= 70) return { bg: "#dbeafe", border: "#bfdbfe", color: "#1d4ed8", bar: "#2563eb" }; if (p >= 50) return { bg: "#fef3c7", border: "#fde68a", color: "#92400e", bar: "#f59e0b" };

    return { bg: "#fee2e2", border: "#fecaca", color: "#991b1b", bar: "#ef4444" };

};

const DOT_COLORS = ["#818cf8", "#f59e0b", "#34d399", "#f87171", "#60a5fa", "#a78bfa", "#fb7185", "#38bdf8",];

const CSS = @page{ size: A4 landscape; margin: 8mm 6mm;}
* { box- sizing: border - box;}
body{ margin: 0; background: #eef3f8; font - family: Arial, Helvetica, sans - serif; color:#0f172a; }
.wrap{ padding: 8px; }
.hero{ background:#0f172a; border - radius: 16px; padding: 18px 22px; margin - bottom: 13px; display: flex; justify - content: space - between; align - items: flex - end; }
.hb{ font - size: 9px; font - weight: 800; letter - spacing: .16em; text - transform: uppercase; color:#93c5fd; margin - bottom: 4px; }
.ht{ font - size: 22px; font - weight: 900; color: #fff; letter - spacing: -.03em; }
.hs{ font - size: 10px; color:#64748b; margin - top: 3px; }
.hd{ font - size: 10px; color:#94a3b8; text - align: right; line - height: 1.7; }
.gkpi{ background: #fff; border: 1px solid #e2e8f0; border - radius: 13px; padding: 12px 14px; margin - bottom: 13px; }
.gkpi - h{ font - size: 8px; font - weight: 800; letter - spacing: .1em; text - transform: uppercase; color:#64748b; margin - bottom: 9px; }
.krow{ display: flex; gap: 7px; flex - wrap: wrap; }
.kp{ flex: 1; min - width: 78px; border - radius: 9px; padding: 9px 10px; text - align: center; }
.kp.l{ font - size: 7px; font - weight: 700; letter - spacing: .07em; text - transform: uppercase; margin - bottom: 4px; }
.kp.v{ font - size: 20px; font - weight: 900; line - height: 1; letter - spacing: -.03em; }
.kp.s{ font - size: 7.5px; margin - top: 3px; }
.kp - b{ background: #eff6ff; border: 1px solid #bfdbfe; }.kp - b.l,.kp - b.v{ color:#1d4ed8; }
.kp - g{ background: #f0fdf4; border: 1px solid #bbf7d0; }.kp - g.l,.kp - g.v{ color:#166534; }
.kp - r{ background: #fff1f2; border: 1px solid #fecdd3; }.kp - r.l,.kp - r.v{ color:#991b1b; }
.kp - a{ background: #fff7ed; border: 1px solid #fed7aa; }.kp - a.l,.kp - a.v{ color:#9a3412; }
.kp - p{ background: #faf5ff; border: 1px solid #e9d5ff; }.kp - p.l,.kp - p.v{ color:#6b21a8; }
.kp - n{ background: #f8fafc; border: 1px solid #e2e8f0; }.kp - n.l,.kp - n.v{ color:#94a3b8; }
.pkp{ flex: 1.3; min - width: 88px; border - radius: 9px; padding: 9px 12px; text - align: center; }
.pkp.l{ font - size: 7px; font - weight: 700; letter - spacing: .07em; text - transform: uppercase; margin - bottom: 4px; }
.pkp.v{ font - size: 20px; font - weight: 900; line - height: 1; }
.bw2{ margin - top: 5px; height: 4px; background: rgba(0, 0, 0, .09); border - radius: 2px; overflow: hidden; }
.br{ height: 100 %; border - radius: 2px; }
.pdiv{ display: flex; align - items: center; gap: 9px; margin: 15px 0 9px; }
.pdiv - line{ flex: 1; height: 1px; background: #e2e8f0; }
.pdiv - badge{ display: flex; align - items: center; gap: 5px; background:#0f172a; color: #e2e8f0; border - radius: 7px; padding: 4px 11px; font - size: 8.5px; font - weight: 800; letter - spacing: .06em; text - transform: uppercase; white - space: nowrap; }
.pdiv - dot{ width: 6px; height: 6px; border - radius: 50 %; flex - shrink: 0; }
.pozet{ background: #fff; border: 1px solid #e2e8f0; border - radius: 11px; padding: 10px 12px; margin - bottom: 9px; display: flex; gap: 7px; align - items: stretch; }
.ps{ flex: 1; border - radius: 7px; padding: 7px 9px; text - align: center; }
.ps.l{ font - size: 7px; font - weight: 700; letter - spacing: .07em; text - transform: uppercase; margin - bottom: 3px; }
.ps.v{ font - size: 16px; font - weight: 900; line - height: 1; letter - spacing: -.02em; }
.psperf{ flex: 1.4; border - radius: 7px; padding: 7px 11px; display: flex; align - items: center; gap: 8px; }
.psperf.lp{ font - size: 7px; font - weight: 700; letter - spacing: .07em; text - transform: uppercase; margin - bottom: 2px; }
.psperf.vp{ font - size: 18px; font - weight: 900; line - height: 1; }
.psperf - bar{ flex: 1; }
.stcard{ background: #fff; border: 1px solid #e2e8f0; border - radius: 11px; overflow: visible; margin - bottom: 13px; }
.sth{ padding: 7px 13px; background: #f8fafc; border - bottom: 1px solid #e2e8f0; display: flex; align - items: center; justify - content: space - between; }
.sth span{ font - size: 8px; font - weight: 800; letter - spacing: .1em; text - transform: uppercase; color:#64748b; }
.scnt{ font - size: 8px; background: #e2e8f0; color:#475569; border - radius: 999px; padding: 2px 8px; font - weight: 700; }
table{ width: 100 %; border - collapse: collapse; table - layout: fixed; }
thead th{ background:#1e293b; color: #cbd5e1; font - size: 7.5px; font - weight: 700; padding: 7px 7px; text - align: left; text - transform: uppercase; letter - spacing: .05em; border - right: 1px solid #334155; }
thead th: last - child{ border - right: none; }
tbody td{ padding: 6px 7px; font - size: 8.5px; border - bottom: 1px solid #f1f5f9; border - right: 1px solid #f1f5f9; vertical - align: middle; color:#1e293b; word -break: break-word; }
tbody td: last - child{ border - right: none; }
tbody tr: last - child td{ border - bottom: none; }
tbody tr: nth - child(even) td{ background: #f8fafc; }
.sno{ font - weight: 800; color:#1e3a8a; }
.ptm{ font - weight: 700; color:#0f172a; }
.pts{ font - size: 7.5px; color:#94a3b8; margin - top: 2px; }
.badge{ display: inline - flex; align - items: center; gap: 3px; padding: 2px 7px; border - radius: 999px; font - size: 7.5px; font - weight: 800; white - space: nowrap; }
.badge::before{ content: ''; width: 5px; height: 5px; border - radius: 50 %; display: inline - block; flex - shrink: 0; }
.bs{ background: #dcfce7; color:#166534; }.bs::before{ background:#16a34a; }
.bw{ background: #fff7ed; color:#9a3412; }.bw::before{ background: #ea580c; }
.bd{ background: #fee2e2; color:#991b1b; }.bd::before{ background: #dc2626; }
.bn{ background: #f1f5f9; color:#475569; }.bn::before{ background:#64748b; }
.fn{ color: #dc2626; font - weight: 800; }
.fp{ color:#15803d; font - weight: 800; }
.fz{ color:#94a3b8; }
.footer{ margin - top: 8px; display: flex; justify - content: space - between; font - size: 8px; color:#94a3b8; padding: 0 2px; }
.pdiv,.pozet,.sth{ break-inside: avoid; page -break-inside: avoid; }
.stcard{ break-inside: auto; page -break-inside: auto; }
thead{ display: table - header - group; }
tfoot{ display: table - footer - group; }
tr{ break-inside: avoid; page -break-inside: avoid; }
tbody td{ break-inside: avoid; page -break-inside: avoid; };

const buildGlobalKpi = (summaries, bolge) => {
    const t = summaries.reduce((a, s) => { a.talep += Number(s.talep || 0); a.tedarik += Number(s.tedarik || 0); a.edilmeyen += Number(s.edilmeyen || 0); a.gec += Number(s.gec_tedarik || 0); a.sho_b += Number(s.sho_basilan || 0); a.spot += Number(s.spot || 0); a.filo += Number(s.filo || 0); return a; }, { talep: 0, tedarik: 0, edilmeyen: 0, gec: 0, sho_b: 0, spot: 0, filo: 0, });

    const zam = Math.max(0, t.tedarik - t.gec);
    const perf = t.talep > 0 ? Math.max(0, Math.min(100, Math.round((zam / t.talep) * 100))) : 0;
    const sho = t.tedarik > 0 ? Math.round((t.sho_b / t.tedarik) * 100) : 0;
    const ps = perfStyle(perf);

    return `

const buildProjeBlock = (summary, rows, dotColor) => {const ozet = hesaplaOzet(summary, rows);

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
        ? `< div class="ps kp-r" ><div class="l">Edilmeyen</div><div class="v">${ed}</div></div > `
        : `< div class="ps kp-n" ><div class="l">Edilmeyen</div><div class="v">&mdash;</div></div > `;

const gecHtml =
    gec > 0
        ? `< div class="ps kp-a" ><div class="l">Ge&ccedil; Tedarik</div><div class="v">${gec}</div></div > `
        : `< div class="ps kp-n" ><div class="l">Ge&ccedil; Tedarik</div><div class="v">&mdash;</div></div > `;

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

        let farkHtml = `< span class="fz" >& mdash;</span > `;

        if (fark !== null && fark !== undefined && fark !== "-") {
            const n = parseFloat(fark);

            if (!isNaN(n)) {
                farkHtml =
                    n < 0
                        ? `< span class="fn" > ${ n }</span > `
                        : n > 0
                            ? `< span class="fp" > +${ n }</span > `
                            : `< span class="fz" > 0</span > `;
            }
        }

        return `

    return `

const buildHtml = (summaries, data, bolge) => {const guncelSummaries = summaries.map((s) => {const key = String(s.proje || "").trim().toLowerCase();const rows = data.filter((r) => String(r.proje || "").trim().toLowerCase() === key);return hesaplaOzet(s, rows);});

const bloklar = guncelSummaries
    .map((s, i) => {
        const key = String(s.proje || "").trim().toLowerCase();
        const rows = data.filter((r) => String(r.proje || "").trim().toLowerCase() === key);
        return buildProjeBlock(s, rows, DOT_COLORS[i % DOT_COLORS.length]);
    })
    .join("\n");

return `< !DOCTYPE html >

const buildExcelBuffer = ({ item, bolge }) => {
        const { data = [], summaries = [] } = item || {};

        const summaryRows = summaries.map((s) => ({
            Bölge: s.bolge || bolge || "-",
            Proje: s.proje || "-",
            Talep: Number(s.talep || 0),
            Tedarik: Number(s.tedarik || 0),
            "Tedarik Edilmeyen": Number(s.edilmeyen || 0),
            Filo: Number(s.filo || 0),
            Spot: Number(s.spot || 0),
            "SHÖ Basılan": Number(s.sho_basilan || 0),
            "SHÖ Basılmayan": Number(s.sho_basilmayan || 0),
            "Geç Tedarik": Number(s.gec_tedarik || 0),
            Oran: `%${s.oran || 0}`,
        }));

        const sonrakiIsGunuSaat6 = (v) => {
            const d = new Date(v);
            if (isNaN(d.getTime())) return null;

            const limit = new Date(d);
            limit.setDate(limit.getDate() + 1);

            while (limit.getDay() === 0 || limit.getDay() === 6) {
                limit.setDate(limit.getDate() + 1);
            }

            limit.setHours(6, 0, 0, 0);
            return limit;
        };

        const hesaplaTedarikDurumuExcel = (r) => {
            const seferNo = String(r?.seferNo || "")
                .trim()
                .toLocaleLowerCase("tr-TR");

            const yuklemeyeGelis = String(r?.yuklemeyeGelis || "").trim();

            if (
                !yuklemeyeGelis ||
                yuklemeyeGelis === "-" ||
                seferNo.includes("planlamada") ||
                seferNo.includes("planlanmadı")
            ) {
                return "Tedarik Edilemeyen";
            }

            const yukleme = r?.yuklemeTarihi ? new Date(r.yuklemeTarihi) : null;
            const gelis = r?.yuklemeyeGelis ? new Date(r.yuklemeyeGelis) : null;

            if (!yukleme || !gelis) return "-";
            if (isNaN(yukleme.getTime()) || isNaN(gelis.getTime())) return "-";

            const limit = sonrakiIsGunuSaat6(yukleme);

            return gelis > limit ? "Geç Tedarik" : "-";
        };

        const detailRows = data.map((r) => ({
            Bölge: r.bolge || "-",
            Proje: r.proje || "-",
            "Sefer No": r.seferNo || "-",
            "Talep No": r.talepNo || "-",
            Müşteri: r.musteri || "-",
            "Yükleme İl": r.yuklemeIl || "-",
            "Yükleme İlçe": r.yuklemeIlce || "-",
            "Yükleme Noktası": r.yuklemeNoktasi || getYukleme(r),
            "Teslim İl": r.teslimIl || "-",
            "Teslim İlçe": r.teslimIlce || "-",
            "Teslim Noktası": r.teslimNoktasi || getTeslim(r),
            "Yükleme Tarihi": r.yuklemeTarihi || "-",
            "Yüklemeye Geliş": r.yuklemeyeGelis || "-",
            "Fark Saat": r.farkSaat ?? "-",
            "Tedarik Durumu": hesaplaTedarikDurumuExcel(r),
            Durum: r.durumText || statusText(r),
            "Araç Tipi": r.aracTipi || "-",
        }));

        const wb = XLSX.utils.book_new();

        const wsSummary = XLSX.utils.json_to_sheet(
            summaryRows.length > 0
                ? summaryRows
                : [{ Bilgi: "Özet verisi bulunamadı." }]
        );

        const wsDetail = XLSX.utils.json_to_sheet(
            detailRows.length > 0
                ? detailRows
                : [{ Bilgi: "Detay sefer verisi bulunamadı." }]
        );

        wsSummary["!cols"] = [
            { wch: 16 },
            { wch: 32 },
            { wch: 10 },
            { wch: 10 },
            { wch: 18 },
            { wch: 10 },
            { wch: 10 },
            { wch: 14 },
            { wch: 16 },
            { wch: 14 },
            { wch: 10 },
        ];

        wsDetail["!cols"] = [
            { wch: 14 },
            { wch: 28 },
            { wch: 18 },
            { wch: 18 },
            { wch: 36 },
            { wch: 14 },
            { wch: 16 },
            { wch: 34 },
            { wch: 14 },
            { wch: 16 },
            { wch: 34 },
            { wch: 20 },
            { wch: 20 },
            { wch: 10 },
            { wch: 22 },
            { wch: 18 },
            { wch: 20 },
        ];

        const headerStyle = {
            font: { bold: true, color: { rgb: "FFFFFF" }, sz: 12 },
            fill: { fgColor: { rgb: "111827" } },
            alignment: { horizontal: "center", vertical: "center" },
            border: {
                top: { style: "thin", color: { rgb: "334155" } },
                bottom: { style: "thin", color: { rgb: "334155" } },
                left: { style: "thin", color: { rgb: "334155" } },
                right: { style: "thin", color: { rgb: "334155" } },
            },
        };

        const bodyStyle = {
            font: { color: { rgb: "0F172A" }, sz: 10 },
            alignment: { vertical: "center", wrapText: true },
            border: {
                bottom: { style: "thin", color: { rgb: "E5E7EB" } },
            },
        };

        const applyModernStyle = (ws) => {
            if (!ws["!ref"]) return;

            const range = XLSX.utils.decode_range(ws["!ref"]);

            for (let C = range.s.c; C <= range.e.c; C++) {
                const cellRef = XLSX.utils.encode_cell({ r: 0, c: C });
                if (ws[cellRef]) ws[cellRef].s = headerStyle;
            }

            for (let R = 1; R <= range.e.r; R++) {
                for (let C = range.s.c; C <= range.e.c; C++) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    if (!ws[cellRef]) continue;

                    ws[cellRef].s = {
                        ...bodyStyle,
                        fill: {
                            fgColor: { rgb: R % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                        },
                    };
                }
            }
        };

        applyModernStyle(wsSummary);
        applyModernStyle(wsDetail);

        const tedarikDurumuColIndex = detailRows.length
            ? Object.keys(detailRows[0]).indexOf("Tedarik Durumu")
            : -1;

        if (tedarikDurumuColIndex >= 0 && wsDetail["!ref"]) {
            const range = XLSX.utils.decode_range(wsDetail["!ref"]);

            for (let R = 1; R <= range.e.r; R++) {
                const cellRef = XLSX.utils.encode_cell({
                    r: R,
                    c: tedarikDurumuColIndex,
                });

                const cell = wsDetail[cellRef];

                if (cell?.v === "Geç Tedarik") {
                    cell.s = {
                        ...cell.s,
                        font: {
                            bold: true,
                            color: { rgb: "991B1B" },
                            sz: 10,
                        },
                        fill: {
                            fgColor: { rgb: "FEE2E2" },
                        },
                        alignment: {
                            horizontal: "center",
                            vertical: "center",
                            wrapText: true,
                        },
                    };
                }

                if (cell?.v === "Tedarik Edilemeyen") {
                    cell.s = {
                        ...cell.s,
                        font: {
                            bold: true,
                            color: { rgb: "92400E" },
                            sz: 10,
                        },
                        fill: {
                            fgColor: { rgb: "FEF3C7" },
                        },
                        alignment: {
                            horizontal: "center",
                            vertical: "center",
                            wrapText: true,
                        },
                    };
                }
            }
        }

        wsSummary["!autofilter"] = { ref: wsSummary["!ref"] };
        wsDetail["!autofilter"] = { ref: wsDetail["!ref"] };

        wsSummary["!freeze"] = { xSplit: 0, ySplit: 1 };
        wsDetail["!freeze"] = { xSplit: 0, ySplit: 1 };

        wsSummary["!rows"] = [{ hpt: 28 }];
        wsDetail["!rows"] = [{ hpt: 28 }];

        XLSX.utils.book_append_sheet(wb, wsSummary, "Özet");
        XLSX.utils.book_append_sheet(wb, wsDetail, "Detay Seferler");

        return XLSX.write(wb, {
            type: "buffer",
            bookType: "xlsx",
        });
    };

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
                    results.push({
                        email: "-",
                        ok: false,
                        message: "Email bulunamadı.",
                    });
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
                        "--disable-gpu"
                    ]
                });
                const page = await browser.newPage();

                await page.setContent(html, {
                    waitUntil: "domcontentloaded",
                    timeout: 120000,
                });

                await page.emulateMediaType("screen");
                await new Promise((resolve) => setTimeout(resolve, 500));

                const pdfBuffer = await page.pdf({
                    format: "A4",
                    landscape: true,
                    printBackground: true,
                    margin: {
                        top: "8mm",
                        right: "6mm",
                        bottom: "8mm",
                        left: "6mm",
                    },
                });

                await browser.close();
                browser = null;

                console.log("📨 SMTP ile mail gönderiliyor:", email);

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

                results.push({
                    email,
                    ok: true,
                    messageId: info.messageId,
                });
            }

            return results;
        } catch (err) {
            if (browser) {
                await browser.close().catch(() => { });
            }

            throw err;
        }

    };

    /* =======================TMS ORDERS HELPERS======================= */

    const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

    const cache = new Map(); const CACHE_TTL_MS = 5 * 60 * 1000;

    function cacheKey({ startDate, endDate, userId }) { return ${ userId }| ${ startDate }| ${ endDate }; }

    async function fetchWithTimeout(url, options, timeoutMs) {
        const controller = new AbortController(); const t = setTimeout(() => controller.abort(), timeoutMs);

        try {
            return await fetch(url, { ...options, signal: controller.signal });
        } finally {
            clearTimeout(t);
        }

    }

    function extractFirstItem(data) { return data?.data?.[0] || data?.Data?.[0] || data?.items?.[0] || data?.[0] || null; }

    async function tmsordersHandler(req, res) { const rid = ${ Date.now() } -${ Math.random().toString(16).slice(2) };

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

                if (attempt < RETRIES) {
                    await sleep(600);
                }
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

/* =======================ROUTES======================= */

app.get("/", (req, res) => { res.send("Tedarik Analiz Backend is running"); });

app.get("/health", (req, res) => { res.json({ ok: true, service: "tedarik-analiz-backend-clean" }); });

app.get("/routes", (req, res) => { res.json({ routes: ["GET /", "GET /health", "GET /routes", "POST /tmsorders", "POST /tmsorders/week", "POST /send-analiz-mail", "POST /download-analiz-pdf", "POST /download-analiz-excel",], allowedOrigins: ALLOWED_ORIGINS, }); });

app.post("/tmsorders", tmsordersHandler); app.post("/tmsorders/week", tmsordersHandler);

app.post("/download-analiz-pdf", async (req, res) => {
    let browser;

    try {
        const { item, bolge } = req.body;

        if (!item) {
            return res.status(400).json({
                ok: false,
                message: "PDF için item verisi bulunamadı.",
            });
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

        await page.setContent(html, {
            waitUntil: "domcontentloaded",
            timeout: 120000,
        });

        await page.emulateMediaType("screen");
        await new Promise((resolve) => setTimeout(resolve, 500));

        const pdfBuffer = await page.pdf({
            format: "A4",
            landscape: true,
            printBackground: true,
            margin: {
                top: "8mm",
                right: "6mm",
                bottom: "8mm",
                left: "6mm",
            },
        });

        await browser.close();
        browser = null;

        const fileName = `sefer-analiz-raporu-${slugify(bolge || "bolge")}.pdf`;

        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);

        return res.send(pdfBuffer);
    } catch (err) {
        if (browser) {
            await browser.close().catch(() => { });
        }

        console.error("PDF oluşturma hatası:", err);

        return res.status(500).json({
            ok: false,
            message: err.message || "PDF oluşturulamadı.",
        });
    }

});

app.post("/download-analiz-excel", async (req, res) => {
    try {
        const { item, bolge } = req.body;

        if (!item) {
            return res.status(400).json({
                ok: false,
                message: "Excel için item verisi bulunamadı.",
            });
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

        return res.status(500).json({
            ok: false,
            message: err.message || "Excel oluşturulamadı.",
        });
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

        const results = await analizMailGonder({
            mailPayload,
            bolge,
        });

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

/* =======================CRON======================= */

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

cron.schedule("0 10,17 * * *", async () => { await otomatikAnalizMailGonder(); }, { timezone: "Europe/Istanbul", });

/* =======================START======================= */

app.listen(PORT, () => { console.log("🚀 Server listening on port", PORT); });