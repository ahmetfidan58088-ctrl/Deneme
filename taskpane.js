// Office.js hazır
Office.onReady((info) => {
    console.log("Office.js hazır. Host:", info.host);
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (analyzeBtn) analyzeBtn.addEventListener("click", analyzeData);
    else console.error("Buton bulunamadı!");
});

async function analyzeData() {
    console.log("Analiz başlıyor...");
    showLoading(true);
    hideResult();
    hideError();

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load("values, rowCount, columnCount, address");
            sheet.load("name");
            await context.sync();

            if (!usedRange.values || usedRange.values.length < 2) {
                throw new Error("Veri bulunamadı! En az bir başlık satırı ve bir veri satırı olmalı.");
            }

            const headers = usedRange.values[0].map(h => String(h || "").trim());
            const dataRows = usedRange.values.slice(1);

            // Sütun indekslerini bul (genişletilmiş alias listesi)
            const colIndex = {
                date: findColumn(headers, ["tarih", "date", "faturalama", "islem_tarihi"]),
                product: findColumn(headers, ["urun", "ürün", "product", "model"]),
                channel: findColumn(headers, ["kanal", "bayi", "channel", "dealer", "müşteri"]),
                quantity: findColumn(headers, ["adet", "miktar", "quantity"]),
                revenue: findColumn(headers, ["tutar", "ciro", "revenue", "satış_tutarı"]),
                stock: findColumn(headers, ["stok", "stock", "inventory"]),
                cost: findColumn(headers, ["maliyet", "cost", "gider"]),
                budget: findColumn(headers, ["bütçe", "butce", "budget"]),
                actual: findColumn(headers, ["gerçekleşen", "actual"]),
                // İnşaat dashboard için ek sütunlar
                region: findColumn(headers, ["bölge", "region", "bolge"]),
                status: findColumn(headers, ["durum", "status", "state"]),
                phase: findColumn(headers, ["faz", "phase", "aşama"]),
                projectType: findColumn(headers, ["proje tipi", "project type", "tip"]),
                contractor: findColumn(headers, ["yüklenici", "contractor", "firma"]),
                department: findColumn(headers, ["departman", "department"]),
                safetyIncidents: findColumn(headers, ["güvenlik", "safety", "incident", "olay"]),
            };

            // Filtre değerlerini al
            const startDateStr = document.getElementById("dateStart").value;
            const endDateStr = document.getElementById("dateEnd").value;
            const selectedChannel = document.getElementById("channelFilter").value;
            const selectedProduct = document.getElementById("productFilter").value;

            // Filtreleme
            let filteredRows = dataRows;
            if (colIndex.date !== -1 && (startDateStr || endDateStr)) {
                filteredRows = filteredRows.filter(row => {
                    const cellDate = parseDate(row[colIndex.date]);
                    if (!cellDate) return true;
                    const start = startDateStr ? parseDate(startDateStr) : null;
                    const end = endDateStr ? parseDate(endDateStr) : null;
                    if (start && cellDate < start) return false;
                    if (end && cellDate > end) return false;
                    return true;
                });
            }
            if (colIndex.channel !== -1 && selectedChannel) {
                filteredRows = filteredRows.filter(row => String(row[colIndex.channel] || "").trim() === selectedChannel);
            }
            if (colIndex.product !== -1 && selectedProduct) {
                filteredRows = filteredRows.filter(row => String(row[colIndex.product] || "").trim() === selectedProduct);
            }

            // Metrikler
            const metrics = calculateMetrics(filteredRows, colIndex);
            const topProducts = getTopProducts(filteredRows, colIndex);
            const dealerPerformance = getDealerPerformance(filteredRows, colIndex);
            const stockRisk = getStockRisk(filteredRows, colIndex);
            const financeSummary = getFinanceSummary(filteredRows, colIndex);
            
            // İnşaat verilerini hazırla
            const constructionData = prepareConstructionData(filteredRows, colIndex);

            // Dashboard sayfaları
            await createExecutiveDashboard(context, metrics, topProducts, dealerPerformance);
            await createSalesDashboard(context, metrics, topProducts);
            await createStockDashboard(context, stockRisk);
            await createFinanceDashboard(context, financeSummary);
            await createDealerDashboard(context, dealerPerformance);
            
            // Yeni: İnşaat dashboard (eğer veri varsa)
            if (constructionData.hasData) {
                await createConstructionDashboard(context, constructionData);
            }

            // Filtre dropdown'larını güncelle
            updateFilterDropdowns(headers, dataRows, colIndex);

            // Task pane özeti
            let resultText = `✅ Analiz tamamlandı!\n\n📄 Sayfa: ${sheet.name}\n📍 Aralık: ${usedRange.address}\n📊 Satır: ${usedRange.rowCount}\n📈 Sütun: ${usedRange.columnCount}\n\n`;
            resultText += `🔍 Başlıklar: ${headers.join(", ")}\n\n`;
            resultText += `📊 Toplam Adet: ${metrics.totalQuantity}\n💰 Toplam Ciro: ${metrics.totalRevenue.toLocaleString()} TL\n`;
            resultText += `📉 Ortalama Adet: ${metrics.avgQuantity.toFixed(2)}\n🏆 En Çok Satan Ürün: ${topProducts[0]?.product || "-"} (${topProducts[0]?.quantity || 0} adet)\n`;
            resultText += `⚠️ Aykırı Değer Sayısı: ${metrics.outliers}\n\n`;
            resultText += `📌 Dashboard sayfaları oluşturuldu: EXECUTIVE_DASHBOARD, SALES_DASHBOARD, STOCK_DASHBOARD, FINANCE_DASHBOARD, DEALER_DASHBOARD`;
            if (constructionData.hasData) {
                resultText += `\n🏗️ CONSTRUCTION_DASHBOARD eklendi (bölge, durum, faz analizleri).`;
            }

            showResult(resultText);
            await context.sync();
        });
    } catch (error) {
        console.error("Hata:", error);
        showError("Hata: " + error.message);
    } finally {
        showLoading(false);
    }
}

// ========== YARDIMCI FONKSİYONLAR ==========
function findColumn(headers, candidates) {
    for (let i = 0; i < headers.length; i++) {
        const h = headers[i].toLowerCase();
        if (candidates.some(c => h.includes(c))) return i;
    }
    return -1;
}

function parseDate(value) {
    if (!value) return null;
    if (value instanceof Date) return value;
    const str = String(value);
    let day, month, year;
    if (str.includes(".")) {
        [day, month, year] = str.split(".");
    } else if (str.includes("-")) {
        [year, month, day] = str.split("-");
    } else {
        return null;
    }
    const d = new Date(year, month-1, day);
    return isNaN(d.getTime()) ? null : d;
}

function calculateMetrics(rows, colIndex) {
    let totalQuantity = 0, totalRevenue = 0, quantities = [];
    for (const row of rows) {
        if (colIndex.quantity !== -1) {
            const qty = parseFloat(row[colIndex.quantity]);
            if (!isNaN(qty)) {
                totalQuantity += qty;
                quantities.push(qty);
            }
        }
        if (colIndex.revenue !== -1) {
            const rev = parseFloat(row[colIndex.revenue]);
            if (!isNaN(rev)) totalRevenue += rev;
        }
    }
    const avgQuantity = quantities.length ? totalQuantity / quantities.length : 0;
    let outliers = 0;
    if (quantities.length) {
        const mean = avgQuantity;
        const variance = quantities.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0) / quantities.length;
        const std = Math.sqrt(variance);
        const threshold = 2 * std;
        outliers = quantities.filter(v => Math.abs(v - mean) > threshold).length;
    }
    return { totalQuantity, totalRevenue, avgQuantity, outliers };
}

function getTopProducts(rows, colIndex, topN = 5) {
    if (colIndex.product === -1 || colIndex.quantity === -1) return [];
    const productMap = new Map();
    for (const row of rows) {
        const prod = String(row[colIndex.product] || "").trim();
        if (!prod) continue;
        const qty = parseFloat(row[colIndex.quantity]);
        if (!isNaN(qty)) productMap.set(prod, (productMap.get(prod) || 0) + qty);
    }
    return Array.from(productMap.entries())
        .map(([product, quantity]) => ({ product, quantity }))
        .sort((a,b) => b.quantity - a.quantity)
        .slice(0, topN);
}

function getDealerPerformance(rows, colIndex) {
    if (colIndex.channel === -1 || colIndex.quantity === -1) return [];
    const dealerMap = new Map();
    for (const row of rows) {
        const dealer = String(row[colIndex.channel] || "").trim();
        if (!dealer) continue;
        const qty = parseFloat(row[colIndex.quantity]);
        if (!isNaN(qty)) dealerMap.set(dealer, (dealerMap.get(dealer) || 0) + qty);
    }
    return Array.from(dealerMap.entries())
        .map(([dealer, quantity]) => ({ dealer, quantity }))
        .sort((a,b) => b.quantity - a.quantity);
}

function getStockRisk(rows, colIndex) {
    if (colIndex.stock === -1 || colIndex.product === -1) return [];
    const riskMap = [];
    for (const row of rows) {
        const product = String(row[colIndex.product] || "").trim();
        const stock = parseFloat(row[colIndex.stock]);
        if (product && !isNaN(stock)) {
            riskMap.push({ product, stock, risk: stock < 20 ? "critical" : (stock < 50 ? "low" : "ok") });
        }
    }
    return riskMap.filter(r => r.risk !== "ok").sort((a,b) => a.stock - b.stock);
}

function getFinanceSummary(rows, colIndex) {
    if (colIndex.budget === -1 || colIndex.actual === -1) return { totalBudget: 0, totalActual: 0, variance: 0 };
    let totalBudget = 0, totalActual = 0;
    for (const row of rows) {
        const b = parseFloat(row[colIndex.budget]);
        const a = parseFloat(row[colIndex.actual]);
        if (!isNaN(b)) totalBudget += b;
        if (!isNaN(a)) totalActual += a;
    }
    return { totalBudget, totalActual, variance: totalActual - totalBudget };
}

// === İnşaat verilerini hazırlama ===
function prepareConstructionData(rows, colIndex) {
    const result = {
        hasData: false,
        totalProjects: rows.length,
        totalBudget: 0,
        totalCost: 0,
        totalSafetyIncidents: 0,
        regionStats: new Map(),   // region -> { count, budget, cost }
        statusStats: new Map(),
        phaseStats: new Map(),
        projectTypeStats: new Map(),
        contractorStats: new Map(),
        departmentStats: new Map(),
    };

    let budgetFound = false, costFound = false, safetyFound = false;

    for (const row of rows) {
        let budget = 0, cost = 0, safety = 0;

        if (colIndex.budget !== -1) {
            const b = parseFloat(row[colIndex.budget]);
            if (!isNaN(b)) { budget = b; budgetFound = true; }
        }
        if (colIndex.cost !== -1) {
            const c = parseFloat(row[colIndex.cost]);
            if (!isNaN(c)) { cost = c; costFound = true; }
        }
        if (colIndex.safetyIncidents !== -1) {
            const s = parseFloat(row[colIndex.safetyIncidents]);
            if (!isNaN(s)) { safety = s; safetyFound = true; }
        }

        result.totalBudget += budget;
        result.totalCost += cost;
        result.totalSafetyIncidents += safety;

        // Bölge
        if (colIndex.region !== -1) {
            const region = String(row[colIndex.region] || "").trim();
            if (region) {
                const stats = result.regionStats.get(region) || { count: 0, budget: 0, cost: 0 };
                stats.count++;
                stats.budget += budget;
                stats.cost += cost;
                result.regionStats.set(region, stats);
            }
        }

        // Durum
        if (colIndex.status !== -1) {
            const status = String(row[colIndex.status] || "").trim();
            if (status) {
                const stats = result.statusStats.get(status) || { count: 0, budget: 0, cost: 0 };
                stats.count++;
                stats.budget += budget;
                stats.cost += cost;
                result.statusStats.set(status, stats);
            }
        }

        // Faz
        if (colIndex.phase !== -1) {
            const phase = String(row[colIndex.phase] || "").trim();
            if (phase) {
                const stats = result.phaseStats.get(phase) || { count: 0, budget: 0, cost: 0 };
                stats.count++;
                stats.budget += budget;
                stats.cost += cost;
                result.phaseStats.set(phase, stats);
            }
        }

        // Proje tipi
        if (colIndex.projectType !== -1) {
            const type = String(row[colIndex.projectType] || "").trim();
            if (type) {
                const stats = result.projectTypeStats.get(type) || { count: 0, budget: 0, cost: 0 };
                stats.count++;
                stats.budget += budget;
                stats.cost += cost;
                result.projectTypeStats.set(type, stats);
            }
        }

        // Yüklenici
        if (colIndex.contractor !== -1) {
            const con = String(row[colIndex.contractor] || "").trim();
            if (con) {
                const stats = result.contractorStats.get(con) || { count: 0, budget: 0, cost: 0 };
                stats.count++;
                stats.budget += budget;
                stats.cost += cost;
                result.contractorStats.set(con, stats);
            }
        }

        // Departman
        if (colIndex.department !== -1) {
            const dept = String(row[colIndex.department] || "").trim();
            if (dept) {
                const stats = result.departmentStats.get(dept) || { count: 0, budget: 0, cost: 0 };
                stats.count++;
                stats.budget += budget;
                stats.cost += cost;
                result.departmentStats.set(dept, stats);
            }
        }
    }

    result.hasData = (budgetFound || costFound) && (result.regionStats.size > 0 || result.statusStats.size > 0 || result.phaseStats.size > 0);
    return result;
}

// === İnşaat Dashboard oluşturma ===
async function createConstructionDashboard(context, data) {
    let sheet = getOrCreateSheet(context, "CONSTRUCTION_DASHBOARD");
    await context.sync();

    // Temizlik
    const used = sheet.getUsedRange();
    if (used) used.clear();

    // Başlık
    sheet.getRange("A1").values = [["🏗️ CONSTRUCTION DASHBOARD"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 18;
    sheet.getRange("A1").format.font.color = "#2E6B8F";

    // KPI'lar
    const kpis = [
        ["# of Project", data.totalProjects],
        ["Budget", data.totalBudget],
        ["Cost", data.totalCost],
        ["Cost Per Project", data.totalProjects ? data.totalCost / data.totalProjects : 0],
        ["Safety Incidents", data.totalSafetyIncidents]
    ];
    const kpiRange = sheet.getRangeByIndexes(2, 0, kpis.length, 2);
    kpiRange.values = kpis;
    kpiRange.format.font.bold = true;

    // Bölge dağılımı
    if (data.regionStats.size > 0) {
        sheet.getRange("E1").values = [["# of Project by Region"]];
        const regions = Array.from(data.regionStats.entries());
        const regionData = regions.map(([region, stats]) => [region, stats.count]);
        const regionRange = sheet.getRangeByIndexes(2, 4, regionData.length, 2);
        regionRange.values = regionData;
        // Pasta grafiği
        const chart = sheet.charts.add("pie", regionRange, "auto");
        chart.title.text = "Project by Region";
        chart.legend.position = "right";
    }

    // Durum dağılımı
    if (data.statusStats.size > 0) {
        sheet.getRange("H1").values = [["# of Project by Status"]];
        const statuses = Array.from(data.statusStats.entries());
        const statusData = statuses.map(([status, stats]) => [status, stats.count]);
        const statusRange = sheet.getRangeByIndexes(2, 7, statusData.length, 2);
        statusRange.values = statusData;
        const chart = sheet.charts.add("columnClustered", statusRange, "auto");
        chart.title.text = "Project Status";
    }

    // Faz dağılımı
    if (data.phaseStats.size > 0) {
        sheet.getRange("K1").values = [["# of Project by Phase"]];
        const phases = Array.from(data.phaseStats.entries());
        const phaseData = phases.map(([phase, stats]) => [phase, stats.count]);
        const phaseRange = sheet.getRangeByIndexes(2, 10, phaseData.length, 2);
        phaseRange.values = phaseData;
        const chart = sheet.charts.add("barClustered", phaseRange, "auto");
        chart.title.text = "Project Phase";
    }

    // Proje tipine göre maliyet
    if (data.projectTypeStats.size > 0) {
        sheet.getRange("N1").values = [["Cost Per Project by Type"]];
        const types = Array.from(data.projectTypeStats.entries());
        const typeData = types.map(([type, stats]) => [type, stats.cost / stats.count]);
        const typeRange = sheet.getRangeByIndexes(2, 13, typeData.length, 2);
        typeRange.values = typeData;
        const chart = sheet.charts.add("columnClustered", typeRange, "auto");
        chart.title.text = "Avg Cost by Project Type";
    }

    // Açıklama
    sheet.getRange("A30").values = [["Not: Dashboard, verilerinizdeki bölge, durum, faz, proje tipi sütunlarına göre otomatik oluşturulmuştur."]];
    sheet.getRange("A30").format.font.italic = true;
    sheet.getRange("A30").format.font.size = 9;

    sheet.getRange("A:O").format.autofitColumns();
}

// Diğer dashboard fonksiyonları
async function createExecutiveDashboard(context, metrics, topProducts, dealerPerformance) {
    let sheet = getOrCreateSheet(context, "EXECUTIVE_DASHBOARD");
    await context.sync();
    sheet.getRange("A1").values = [["EXECUTIVE DASHBOARD"]];
    sheet.getRange("A1").format.font.bold = true;
    const kpis = [["Total Revenue (TL)", metrics.totalRevenue], ["Total Quantity", metrics.totalQuantity], ["Avg Quantity", metrics.avgQuantity.toFixed(2)], ["Outliers", metrics.outliers]];
    sheet.getRangeByIndexes(2, 0, kpis.length, 2).values = kpis;
    if (topProducts.length) {
        sheet.getRange("E1").values = [["Top 5 Products"]];
        sheet.getRangeByIndexes(2, 4, topProducts.length, 2).values = topProducts.map(p => [p.product, p.quantity]);
    }
    if (dealerPerformance.length) {
        sheet.getRange("H1").values = [["Dealer Performance"]];
        sheet.getRangeByIndexes(2, 7, dealerPerformance.length, 2).values = dealerPerformance.map(d => [d.dealer, d.quantity]);
        const chart = sheet.charts.add("columnClustered", sheet.getRange("H2").getExtendedRange(dealerPerformance.length, 2), "auto");
        chart.title.text = "Dealer Performance";
    }
    sheet.getRange("A:J").format.autofitColumns();
}

async function createSalesDashboard(context, metrics, topProducts) {
    let sheet = getOrCreateSheet(context, "SALES_DASHBOARD");
    await context.sync();
    sheet.getRange("A1").values = [["SALES DASHBOARD"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A3").values = [["Total Revenue", metrics.totalRevenue]];
    sheet.getRange("A4").values = [["Total Quantity", metrics.totalQuantity]];
    if (topProducts.length) {
        sheet.getRange("A6").values = [["Top Products"]];
        const prodRange = sheet.getRangeByIndexes(6, 0, topProducts.length, 2);
        prodRange.values = topProducts.map(p => [p.product, p.quantity]);
        const chart = sheet.charts.add("columnClustered", sheet.getRange("A7").getExtendedRange(topProducts.length, 2), "auto");
        chart.title.text = "Top Selling Products";
    }
}

async function createStockDashboard(context, stockRisk) {
    let sheet = getOrCreateSheet(context, "STOCK_DASHBOARD");
    await context.sync();
    sheet.getRange("A1").values = [["STOCK DASHBOARD - Riskli Ürünler"]];
    sheet.getRange("A1").format.font.bold = true;
    if (stockRisk.length) {
        sheet.getRangeByIndexes(2, 0, stockRisk.length, 3).values = stockRisk.map(r => [r.product, r.stock, r.risk]);
    } else {
        sheet.getRange("A3").values = [["Riskli stok bulunmamaktadır."]];
    }
    sheet.getRange("A:C").format.autofitColumns();
}

async function createFinanceDashboard(context, finance) {
    let sheet = getOrCreateSheet(context, "FINANCE_DASHBOARD");
    await context.sync();
    sheet.getRange("A1").values = [["FINANCE DASHBOARD"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A3").values = [["Total Budget", finance.totalBudget]];
    sheet.getRange("A4").values = [["Total Actual", finance.totalActual]];
    sheet.getRange("A5").values = [["Variance", finance.variance]];
    if (finance.variance > 0) sheet.getRange("A5").format.font.color = "green";
    else if (finance.variance < 0) sheet.getRange("A5").format.font.color = "red";
}

async function createDealerDashboard(context, dealerPerformance) {
    let sheet = getOrCreateSheet(context, "DEALER_DASHBOARD");
    await context.sync();
    sheet.getRange("A1").values = [["DEALER PERFORMANCE"]];
    sheet.getRange("A1").format.font.bold = true;
    if (dealerPerformance.length) {
        const dealerRange = sheet.getRangeByIndexes(2, 0, dealerPerformance.length, 2);
        dealerRange.values = dealerPerformance.map(d => [d.dealer, d.quantity]);
        const chart = sheet.charts.add("barClustered", sheet.getRange("A3").getExtendedRange(dealerPerformance.length, 2), "auto");
        chart.title.text = "Dealer Performance";
    }
}

function getOrCreateSheet(context, sheetName) {
    const sheets = context.workbook.worksheets;
    const sheet = sheets.getItemOrNullObject(sheetName);
    context.load(sheet, "name");
    return sheet;
}

function updateFilterDropdowns(headers, dataRows, colIndex) {
    if (colIndex.channel !== -1) {
        const channels = [...new Set(dataRows.map(row => String(row[colIndex.channel] || "").trim()).filter(v => v))];
        const channelSelect = document.getElementById("channelFilter");
        if (channelSelect) channelSelect.innerHTML = '<option value="">Tümü</option>' + channels.map(c => `<option value="${c}">${c}</option>`).join("");
    }
    if (colIndex.product !== -1) {
        const products = [...new Set(dataRows.map(row => String(row[colIndex.product] || "").trim()).filter(v => v))];
        const productSelect = document.getElementById("productFilter");
        if (productSelect) productSelect.innerHTML = '<option value="">Tümü</option>' + products.map(p => `<option value="${p}">${p}</option>`).join("");
    }
}

// UI yardımcıları
function showLoading(show) {
    const loading = document.getElementById("loading");
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (loading) loading.classList.toggle("hidden", !show);
    if (analyzeBtn) {
        analyzeBtn.disabled = show;
        analyzeBtn.textContent = show ? "⏳ Analiz Ediliyor..." : "📊 Veriyi Analiz Et (Dashboard Üret)";
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