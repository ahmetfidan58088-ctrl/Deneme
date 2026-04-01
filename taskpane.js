// ========== YARDIMCI: Satır numarasını Excel hücre adresine çevir ==========
function cellAddr(row, col) {
    // row: 0-tabanlı, col: 0-tabanlı → örn. (0,0) = "A1"
    let colStr = "";
    let c = col;
    do {
        colStr = String.fromCharCode(65 + (c % 26)) + colStr;
        c = Math.floor(c / 26) - 1;
    } while (c >= 0);
    return colStr + (row + 1);
}

function rangeAddr(row, col, numRows, numCols) {
    return cellAddr(row, col) + ":" + cellAddr(row + numRows - 1, col + numCols - 1);
}

// ========== ANA MODÜL ==========
Office.onReady((info) => {
    console.log("Office.js ready. Host:", info.host);
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (analyzeBtn) analyzeBtn.addEventListener("click", analyzeAllSheets);
    else console.error("Button not found");
});

async function analyzeAllSheets() {
    showLoading(true);
    hideResult();
    hideError();

    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items");
            await context.sync();

            const allData = [];
            for (let i = 0; i < sheets.items.length; i++) {
                const sheet = sheets.items[i];
                const usedRange = sheet.getUsedRange();
                usedRange.load("values");
                sheet.load("name");
                await context.sync();

                if (usedRange.values && usedRange.values.length > 1) {
                    const headers = usedRange.values[0].map(h => String(h || "").trim());
                    const rows = usedRange.values.slice(1);
                    allData.push({ sheetName: sheet.name, headers, rows });
                }
            }

            if (allData.length === 0) {
                throw new Error("Hiç veri bulunamadı.");
            }

            const columnMapping = detectColumnsAcrossSheets(allData);

            const mergedData = { rows: [], headers: [] };
            for (const data of allData) {
                const mapped = mapDataToColumns(data.rows, data.headers, columnMapping);
                mergedData.rows.push(...mapped.rows);
                if (mergedData.headers.length === 0 && mapped.headers.length) {
                    mergedData.headers = mapped.headers;
                }
            }

            const qualityIssues = runQualityChecks(mergedData, columnMapping);

            await createDashboardSheets(context, mergedData, columnMapping, qualityIssues);

            const resultText = `✅ Analiz tamamlandı!\n\n` +
                `📊 Toplam ${mergedData.rows.length} satır veri işlendi.\n` +
                `🔍 Tespit edilen kolonlar: ${Object.entries(columnMapping).map(([k,v]) => `${k}: ${v || "bulunamadı"}`).join(", ")}\n` +
                `⚠️ ${qualityIssues.length} adet veri kalite sorunu tespit edildi.\n\n` +
                `📌 Dashboard sayfaları oluşturuldu: 00_Executive, 01_Sales, 02_Stock, 03_Finance, 04_Channel, 05_Product, 06_DataQuality`;
            showResult(resultText);

            await context.sync();
        });
    } catch (error) {
        console.error("Hata:", error);
        showError("Analiz sırasında hata: " + error.message);
    } finally {
        showLoading(false);
    }
}

// ========== KOLON TANIMA ==========
const ALIASES = {
    date: ["tarih", "date", "islem_tarihi", "siparis_tarihi", "invoice_date", "month", "ay"],
    product: ["urun", "ürün", "product", "model", "malzeme", "item", "sku", "pn", "part_number"],
    quantity: ["adet", "miktar", "quantity", "qty", "satilan_adet", "satis_adedi", "units"],
    revenue: ["ciro", "revenue", "sales_amount", "tutar", "net_satis"],
    stock: ["stok", "stock", "inventory", "mevcut_stok"],
    budget: ["butce", "bütçe", "budget", "plan"],
    actual: ["gerceklesen", "gerçekleşen", "actual", "realized"],
    cost: ["maliyet", "cost", "gider", "expense"],
    channel: ["kanal", "bayi", "channel", "dealer", "customer", "müşteri"],
    region: ["bolge", "bölge", "region"],
    status: ["durum", "status", "state"],
    phase: ["faz", "phase", "asama"],
    projectType: ["proje tipi", "project type", "tip"],
    safetyIncidents: ["güvenlik", "safety", "incident", "olay"]
};

function normalizeString(s) {
    if (!s) return "";
    return s.toLowerCase()
        .replace(/ç/g, "c").replace(/ğ/g, "g").replace(/ı/g, "i")
        .replace(/ö/g, "o").replace(/ş/g, "s").replace(/ü/g, "u")
        .replace(/[^a-z0-9]/g, " ").trim();
}

function similarityScore(str1, str2) {
    const tokens1 = str1.split(/\s+/);
    const tokens2 = str2.split(/\s+/);
    let match = 0;
    for (let t of tokens1) { if (tokens2.includes(t)) match++; }
    return match / Math.max(tokens1.length, tokens2.length);
}

function detectColumnsAcrossSheets(allData) {
    const mapping = {};
    for (let canonical of Object.keys(ALIASES)) { mapping[canonical] = null; }

    const allHeaders = new Set();
    for (const data of allData) { for (const h of data.headers) { allHeaders.add(h); } }

    for (let [canonical, aliases] of Object.entries(ALIASES)) {
        let bestHeader = null, bestScore = 0;
        for (const header of allHeaders) {
            const normHeader = normalizeString(header);
            for (const alias of aliases) {
                const score = similarityScore(normHeader, normalizeString(alias));
                if (score > bestScore) { bestScore = score; bestHeader = header; }
            }
        }
        if (bestScore > 0.6) { mapping[canonical] = bestHeader; }
    }
    return mapping;
}

function mapDataToColumns(rows, headers, mapping) {
    const colIndex = {};
    for (let [canonical, header] of Object.entries(mapping)) {
        colIndex[canonical] = header ? headers.indexOf(header) : -1;
    }

    const mappedRows = [];
    for (const row of rows) {
        const mapped = {};
        for (let [canonical, idx] of Object.entries(colIndex)) {
            if (idx !== -1 && idx < row.length) {
                let val = row[idx];
                if (canonical === "date") val = parseDate(val);
                else if (["quantity", "revenue", "stock", "budget", "actual", "cost"].includes(canonical)) val = parseNumber(val);
                else val = val !== undefined && val !== null ? String(val).trim() : "";
                mapped[canonical] = val;
            } else {
                mapped[canonical] = canonical === "date" ? null : "";
            }
        }
        mappedRows.push(mapped);
    }
    return { rows: mappedRows, headers: Object.keys(mapping) };
}

function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    const str = String(val);
    let day, month, year;
    if (str.includes(".")) { [day, month, year] = str.split("."); }
    else if (str.includes("-")) { [year, month, day] = str.split("-"); }
    else { return null; }
    const d = new Date(year, month - 1, day);
    return isNaN(d.getTime()) ? null : d;
}

function parseNumber(val) {
    if (val === undefined || val === null) return NaN;
    if (typeof val === "number") return val;
    const s = String(val).replace(/[^0-9,\.\-]/g, "").replace(",", ".");
    const n = parseFloat(s);
    return isNaN(n) ? NaN : n;
}

// ========== VERİ KALİTE KONTROLLERİ ==========
function runQualityChecks(mergedData, mapping) {
    const issues = [];
    const rows = mergedData.rows;
    if (rows.length === 0) return issues;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        for (let col of Object.keys(row)) {
            if (row[col] === undefined || row[col] === null || row[col] === "") {
                issues.push({ sheet: "Tüm veri", row: i + 2, column: col, issue: "Eksik değer", severity: row[col] === null ? "medium" : "low", suggestion: "Hücreyi doldurun veya varsayılan değer atayın." });
            }
        }
    }

    if (mapping.date) {
        for (let i = 0; i < rows.length; i++) {
            const d = rows[i].date;
            if (d === null && rows[i].date !== undefined && rows[i].date !== "") {
                issues.push({ sheet: "Tüm veri", row: i + 2, column: mapping.date, issue: "Geçersiz tarih formatı", severity: "medium", suggestion: "Tarih formatını GG.AA.YYYY veya YYYY-AA-GG olarak düzeltin." });
            }
        }
    }

    const numericCols = ["quantity", "revenue", "stock", "budget", "actual", "cost"];
    for (let col of numericCols) {
        if (mapping[col]) {
            for (let i = 0; i < rows.length; i++) {
                const val = rows[i][col];
                if (val !== undefined && val !== null && val !== "" && isNaN(val)) {
                    issues.push({ sheet: "Tüm veri", row: i + 2, column: mapping[col], issue: "Sayısal olmayan değer", severity: "high", suggestion: "Değeri sayıya çevirin (virgül, TL gibi işaretleri temizleyin)." });
                }
            }
        }
    }

    const seen = new Set();
    for (let i = 0; i < rows.length; i++) {
        const key = JSON.stringify(rows[i]);
        if (seen.has(key)) {
            issues.push({ sheet: "Tüm veri", row: i + 2, column: "tüm sütunlar", issue: "Tamamen kopya satır", severity: "low", suggestion: "Tekrar eden satırı silin." });
        } else { seen.add(key); }
    }

    return issues;
}

// ========== DASHBOARD SAYFALARI OLUŞTURMA ==========
async function createDashboardSheets(context, data, mapping, issues) {
    const sheetNames = ["00_Executive", "01_Sales", "02_Stock", "03_Finance", "04_Channel", "05_Product", "06_DataQuality"];
    for (let name of sheetNames) {
        try {
            const sheet = context.workbook.worksheets.getItemOrNullObject(name);
            sheet.load("isNullObject");
            await context.sync();
            if (!sheet.isNullObject) { sheet.delete(); await context.sync(); }
        } catch (e) { /* sayfa yoksa atla */ }
    }

    await createExecutiveSheet(context, data, mapping, issues);
    await createSalesSheet(context, data, mapping);
    await createStockSheet(context, data, mapping);
    await createFinanceSheet(context, data, mapping);
    await createChannelSheet(context, data, mapping);
    await createProductSheet(context, data, mapping);
    await createQualitySheet(context, issues);
}

async function createExecutiveSheet(context, data, mapping, issues) {
    const sheet = context.workbook.worksheets.add("00_Executive");
    let r = 0;

    sheet.getRange(cellAddr(r, 0)).values = [["EXECUTIVE DASHBOARD - ÖZET"]];
    sheet.getRange(cellAddr(r, 0)).format.font.bold = true;
    r += 2;

    const totalQty = data.rows.reduce((sum, row) => sum + (isNaN(row.quantity) ? 0 : row.quantity), 0);
    const totalRevenue = data.rows.reduce((sum, row) => sum + (isNaN(row.revenue) ? 0 : row.revenue), 0);
    const avgQty = data.rows.length ? totalQty / data.rows.length : 0;

    sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Toplam Adet", totalQty]]; r++;
    sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Toplam Ciro (TL)", totalRevenue]]; r++;
    sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Ortalama Adet", parseFloat(avgQty.toFixed(2))]]; r++;
    sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Kalite Sorunu Sayısı", issues.length]]; r += 2;

    if (mapping.product && mapping.quantity) {
        const prodMap = new Map();
        for (const row of data.rows) {
            if (row.product && !isNaN(row.quantity)) {
                prodMap.set(row.product, (prodMap.get(row.product) || 0) + row.quantity);
            }
        }
        const topProducts = Array.from(prodMap.entries()).sort((a, b) => b[1] - a[1]).slice(0, 5);
        sheet.getRange(cellAddr(r, 0)).values = [["En Çok Satan Ürünler"]]; r++;
        for (let p of topProducts) {
            sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [[p[0], p[1]]]; r++;
        }
        r++;
    }

    sheet.getRange(cellAddr(r, 0)).values = [["Grafik Önerileri"]]; r++;
    sheet.getRange(cellAddr(r, 0)).values = [["• Satış trendi için çizgi grafik"]]; r++;
    sheet.getRange(cellAddr(r, 0)).values = [["• Kanal dağılımı için pasta grafik"]]; r++;

    sheet.getRange("A:C").format.autofitColumns();
    await context.sync();
}

async function createSalesSheet(context, data, mapping) {
    const sheet = context.workbook.worksheets.add("01_Sales");
    let r = 0;

    sheet.getRange(cellAddr(r, 0)).values = [["SATIŞ ANALİZİ"]];
    sheet.getRange(cellAddr(r, 0)).format.font.bold = true;
    r += 2;

    if (mapping.product && mapping.quantity) {
        const prodMap = new Map();
        for (const row of data.rows) {
            if (row.product && !isNaN(row.quantity)) {
                prodMap.set(row.product, (prodMap.get(row.product) || 0) + row.quantity);
            }
        }
        const topProducts = Array.from(prodMap.entries()).sort((a, b) => b[1] - a[1]);

        sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Ürün", "Toplam Adet"]]; r++;
        const dataStartRow = r;
        for (let p of topProducts) {
            sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [[p[0], p[1]]]; r++;
        }

        if (topProducts.length > 0) {
            await context.sync();
            const chartRange = sheet.getRangeByIndexes(dataStartRow, 0, topProducts.length, 2);
            const chart = sheet.charts.add("ColumnClustered", chartRange, "Auto");
            chart.title.text = "Ürün Satış Adetleri";
            chart.legend.position = "Bottom";
        }
    } else {
        sheet.getRange(cellAddr(r, 0)).values = [["Ürün veya adet sütunu bulunamadı."]];
    }

    sheet.getRange("A:C").format.autofitColumns();
    await context.sync();
}

async function createStockSheet(context, data, mapping) {
    const sheet = context.workbook.worksheets.add("02_Stock");
    let r = 0;

    sheet.getRange(cellAddr(r, 0)).values = [["STOK ANALİZİ"]];
    sheet.getRange(cellAddr(r, 0)).format.font.bold = true;
    r += 2;

    if (mapping.product && mapping.stock) {
        const stocks = [];
        for (const row of data.rows) {
            if (row.product && !isNaN(row.stock)) { stocks.push({ product: row.product, stock: row.stock }); }
        }
        const risk = stocks.filter(s => s.stock < 20).sort((a, b) => a.stock - b.stock);
        sheet.getRange(cellAddr(r, 0)).values = [["Kritik Stok (<20)"]]; r++;
        for (let s of risk) {
            sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [[s.product, s.stock]]; r++;
        }
        if (risk.length === 0) {
            sheet.getRange(cellAddr(r, 0)).values = [["Kritik stok seviyesinde ürün bulunamadı."]];
        }
    } else {
        sheet.getRange(cellAddr(r, 0)).values = [["Stok veya ürün sütunu bulunamadı."]];
    }

    sheet.getRange("A:C").format.autofitColumns();
    await context.sync();
}

async function createFinanceSheet(context, data, mapping) {
    const sheet = context.workbook.worksheets.add("03_Finance");
    let r = 0;

    sheet.getRange(cellAddr(r, 0)).values = [["FİNANS ANALİZİ"]];
    sheet.getRange(cellAddr(r, 0)).format.font.bold = true;
    r += 2;

    if (mapping.budget && mapping.actual) {
        let totalBudget = 0, totalActual = 0;
        for (const row of data.rows) {
            if (!isNaN(row.budget)) totalBudget += row.budget;
            if (!isNaN(row.actual)) totalActual += row.actual;
        }
        const variance = totalActual - totalBudget;
        sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Toplam Bütçe", totalBudget]]; r++;
        sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Toplam Gerçekleşen", totalActual]]; r++;
        sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Varyans", variance]];
        sheet.getRange(cellAddr(r, 1)).format.font.color = variance >= 0 ? "#008000" : "#FF0000";
    } else {
        sheet.getRange(cellAddr(r, 0)).values = [["Bütçe veya gerçekleşen sütunu bulunamadı."]];
    }

    sheet.getRange("A:C").format.autofitColumns();
    await context.sync();
}

async function createChannelSheet(context, data, mapping) {
    const sheet = context.workbook.worksheets.add("04_Channel");
    let r = 0;

    sheet.getRange(cellAddr(r, 0)).values = [["KANAL / BAYİ PERFORMANSI"]];
    sheet.getRange(cellAddr(r, 0)).format.font.bold = true;
    r += 2;

    if (mapping.channel && mapping.quantity) {
        const channelMap = new Map();
        for (const row of data.rows) {
            if (row.channel && !isNaN(row.quantity)) {
                channelMap.set(row.channel, (channelMap.get(row.channel) || 0) + row.quantity);
            }
        }
        const channels = Array.from(channelMap.entries()).sort((a, b) => b[1] - a[1]);
        sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Kanal", "Toplam Adet"]]; r++;
        const dataStartRow = r;
        for (let c of channels) {
            sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [[c[0], c[1]]]; r++;
        }

        if (channels.length > 0) {
            await context.sync();
            const chartRange = sheet.getRangeByIndexes(dataStartRow, 0, channels.length, 2);
            const chart = sheet.charts.add("Pie", chartRange, "Auto");
            chart.title.text = "Kanal Dağılımı";
        }
    } else {
        sheet.getRange(cellAddr(r, 0)).values = [["Kanal veya adet sütunu bulunamadı."]];
    }

    sheet.getRange("A:C").format.autofitColumns();
    await context.sync();
}

async function createProductSheet(context, data, mapping) {
    const sheet = context.workbook.worksheets.add("05_Product");
    let r = 0;

    sheet.getRange(cellAddr(r, 0)).values = [["ÜRÜN ANALİZİ"]];
    sheet.getRange(cellAddr(r, 0)).format.font.bold = true;
    r += 2;

    if (mapping.product && mapping.quantity) {
        const prodMap = new Map();
        for (const row of data.rows) {
            if (row.product && !isNaN(row.quantity)) {
                prodMap.set(row.product, (prodMap.get(row.product) || 0) + row.quantity);
            }
        }
        const products = Array.from(prodMap.entries()).sort((a, b) => b[1] - a[1]);
        sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [["Ürün", "Toplam Adet"]]; r++;
        for (let p of products) {
            sheet.getRange(cellAddr(r, 0) + ":" + cellAddr(r, 1)).values = [[p[0], p[1]]]; r++;
        }
    } else {
        sheet.getRange(cellAddr(r, 0)).values = [["Ürün veya adet sütunu bulunamadı."]];
    }

    sheet.getRange("A:C").format.autofitColumns();
    await context.sync();
}

async function createQualitySheet(context, issues) {
    const sheet = context.workbook.worksheets.add("06_DataQuality");
    let r = 0;

    sheet.getRange(cellAddr(r, 0)).values = [["VERİ KALİTE RAPORU"]];
    sheet.getRange(cellAddr(r, 0)).format.font.bold = true;
    r += 2;

    sheet.getRange(rangeAddr(r, 0, 1, 6)).values = [["Sayfa", "Satır", "Sütun", "Sorun", "Şiddet", "Öneri"]];
    sheet.getRange(rangeAddr(r, 0, 1, 6)).format.font.bold = true;
    r++;

    for (let issue of issues) {
        sheet.getRange(rangeAddr(r, 0, 1, 6)).values = [[issue.sheet, issue.row, issue.column, issue.issue, issue.severity, issue.suggestion]];
        r++;
    }

    if (issues.length === 0) {
        sheet.getRange(cellAddr(r, 0)).values = [["✅ Veri kalite sorunu tespit edilmedi."]];
    }

    sheet.getRange("A:F").format.autofitColumns();
    await context.sync();
}

// ========== UI YARDIMCILARI ==========
function showLoading(show) {
    const loading = document.getElementById("loading");
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (loading) loading.classList.toggle("hidden", !show);
    if (analyzeBtn) {
        analyzeBtn.disabled = show;
        analyzeBtn.textContent = show ? "⏳ Analiz Ediliyor..." : "📊 Analiz Başlat";
    }
}
function showResult(text) {
    const resultArea = document.getElementById("resultArea");
    const resultText = document.getElementById("resultText");
    if (resultArea && resultText) {
        resultText.textContent = text;
        resultArea.classList.remove("hidden");
    }
}
function hideResult() { document.getElementById("resultArea")?.classList.add("hidden"); }
function showError(message) {
    const errorArea = document.getElementById("errorArea");
    const errorText = document.getElementById("errorText");
    if (errorArea && errorText) {
        errorText.textContent = message;
        errorArea.classList.remove("hidden");
    }
}
function hideError() { document.getElementById("errorArea")?.classList.add("hidden"); }
