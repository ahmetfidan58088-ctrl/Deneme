Office.onReady((info) => {
    console.log("Office.js ready. Host:", info.host);
    const analyzeBtn = document.getElementById("analyzeBtn");
    if (analyzeBtn) analyzeBtn.addEventListener("click", analyzeData);
    else console.error("Button not found");
});

async function analyzeData() {
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

            // Sütun indekslerini bul
            const colIndex = {
                product: findColumn(headers, ["urun", "ürün", "product", "model"]),
                channel: findColumn(headers, ["kanal", "bayi", "channel", "dealer", "müşteri"]),
                quantity: findColumn(headers, ["adet", "miktar", "quantity"]),
                revenue: findColumn(headers, ["tutar", "ciro", "revenue", "satış_tutarı"]),
                stock: findColumn(headers, ["stok", "stock", "inventory"]),
                cost: findColumn(headers, ["maliyet", "cost", "gider"]),
                budget: findColumn(headers, ["bütçe", "butce", "budget"]),
                actual: findColumn(headers, ["gerçekleşen", "actual"]),
                region: findColumn(headers, ["bölge", "region", "bolge"]),
                status: findColumn(headers, ["durum", "status", "state"]),
                phase: findColumn(headers, ["faz", "phase", "aşama"]),
                projectType: findColumn(headers, ["proje tipi", "project type", "tip"]),
                safetyIncidents: findColumn(headers, ["güvenlik", "safety", "incident", "olay"])
            };

            // Filtre değerlerini al
            const selectedChannel = document.getElementById("channelFilter").value;
            const selectedProduct = document.getElementById("productFilter").value;

            // Filtreleme
            let filteredRows = dataRows;
            if (colIndex.channel !== -1 && selectedChannel) {
                filteredRows = filteredRows.filter(row => String(row[colIndex.channel] || "").trim() === selectedChannel);
            }
            if (colIndex.product !== -1 && selectedProduct) {
                filteredRows = filteredRows.filter(row => String(row[colIndex.product] || "").trim() === selectedProduct);
            }

            // Metrikler ve uyarılar
            const metrics = calculateMetrics(filteredRows, colIndex);
            const warnings = collectWarnings(filteredRows, colIndex, headers);
            const topProducts = getTopProducts(filteredRows, colIndex);
            const channelPerformance = getChannelPerformance(filteredRows, colIndex);
            const stockRisks = getStockRisk(filteredRows, colIndex);
            const financeSummary = getFinanceSummary(filteredRows, colIndex);
            const constructionData = prepareConstructionData(filteredRows, colIndex);

            // Tek dashboard sayfasını oluştur / güncelle
            await createExecutiveDashboard(context, metrics, topProducts, channelPerformance, stockRisks, financeSummary, warnings, constructionData);

            // Filtre dropdown'larını güncelle (kullanıcı sonra seçim yapabilir)
            updateFilterDropdowns(headers, dataRows, colIndex);

            // Task pane özeti
            let resultText = `✅ Analiz tamamlandı!\n\n📄 Sayfa: ${sheet.name}\n📍 Aralık: ${usedRange.address}\n📊 Satır: ${usedRange.rowCount}\n📈 Sütun: ${usedRange.columnCount}\n\n`;
            resultText += `🔍 Başlıklar: ${headers.join(", ")}\n\n`;
            resultText += `📊 Toplam Adet: ${metrics.totalQuantity}\n💰 Toplam Ciro: ${metrics.totalRevenue.toLocaleString()} TL\n`;
            resultText += `📉 Ortalama Adet: ${metrics.avgQuantity.toFixed(2)}\n🏆 En Çok Satın Ürün: ${topProducts[0]?.product || "-"} (${topProducts[0]?.quantity || 0} adet)\n`;
            resultText += `⚠️ Aykırı Değer Sayısı: ${metrics.outliers}\n\n`;
            resultText += `📌 Tüm analizler "EXECUTIVE_DASHBOARD" sayfasında toplanmıştır.`;

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

function collectWarnings(rows, colIndex, headers) {
    const warnings = [];
    if (colIndex.quantity === -1) warnings.push("⚠️ 'Adet' sütunu bulunamadı. Satış adedi analizi yapılamayacak.");
    if (colIndex.revenue === -1) warnings.push("⚠️ 'Ciro/Tutar' sütunu bulunamadı. Gelir analizi yapılamayacak.");
    if (colIndex.channel === -1) warnings.push("⚠️ 'Kanal/Bayi' sütunu bulunamadı. Kanal performansı gösterilemeyecek.");
    if (colIndex.product === -1) warnings.push("⚠️ 'Ürün' sütunu bulunamadı. Ürün bazlı analiz yapılamayacak.");
    if (colIndex.stock === -1) warnings.push("ℹ️ 'Stok' sütunu bulunamadı. Stok riski analizi atlandı.");
    if (colIndex.budget === -1 || colIndex.actual === -1) warnings.push("ℹ️ Bütçe/Gerçekleşen sütunları eksik. Finans analizi sınırlı olacak.");
    return warnings;
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

function getChannelPerformance(rows, colIndex) {
    if (colIndex.channel === -1 || colIndex.quantity === -1) return [];
    const channelMap = new Map();
    for (const row of rows) {
        const channel = String(row[colIndex.channel] || "").trim();
        if (!channel) continue;
        const qty = parseFloat(row[colIndex.quantity]);
        if (!isNaN(qty)) channelMap.set(channel, (channelMap.get(channel) || 0) + qty);
    }
    return Array.from(channelMap.entries())
        .map(([channel, quantity]) => ({ channel, quantity }))
        .sort((a,b) => b.quantity - a.quantity);
}

function getStockRisk(rows, colIndex) {
    if (colIndex.stock === -1 || colIndex.product === -1) return [];
    const riskMap = [];
    for (const row of rows) {
        const product = String(row[colIndex.product] || "").trim();
        const stock = parseFloat(row[colIndex.stock]);
        if (product && !isNaN(stock)) {
            riskMap.push({ product, stock, risk: stock < 20 ? "Kritik" : (stock < 50 ? "Düşük" : "Yeterli") });
        }
    }
    return riskMap.filter(r => r.risk !== "Yeterli").sort((a,b) => a.stock - b.stock);
}

function getFinanceSummary(rows, colIndex) {
    if (colIndex.budget === -1 || colIndex.actual === -1) return null;
    let totalBudget = 0, totalActual = 0;
    for (const row of rows) {
        const b = parseFloat(row[colIndex.budget]);
        const a = parseFloat(row[colIndex.actual]);
        if (!isNaN(b)) totalBudget += b;
        if (!isNaN(a)) totalActual += a;
    }
    return { totalBudget, totalActual, variance: totalActual - totalBudget };
}

function prepareConstructionData(rows, colIndex) {
    const result = { hasData: false, regionStats: new Map(), statusStats: new Map(), phaseStats: new Map(), projectTypeStats: new Map() };
    let budgetFound = false, costFound = false;
    for (const row of rows) {
        let budget = 0, cost = 0;
        if (colIndex.budget !== -1) {
            const b = parseFloat(row[colIndex.budget]);
            if (!isNaN(b)) { budget = b; budgetFound = true; }
        }
        if (colIndex.cost !== -1) {
            const c = parseFloat(row[colIndex.cost]);
            if (!isNaN(c)) { cost = c; costFound = true; }
        }
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
        if (colIndex.status !== -1) {
            const status = String(row[colIndex.status] || "").trim();
            if (status) {
                const stats = result.statusStats.get(status) || { count: 0 };
                stats.count++;
                result.statusStats.set(status, stats);
            }
        }
        if (colIndex.phase !== -1) {
            const phase = String(row[colIndex.phase] || "").trim();
            if (phase) {
                const stats = result.phaseStats.get(phase) || { count: 0 };
                stats.count++;
                result.phaseStats.set(phase, stats);
            }
        }
        if (colIndex.projectType !== -1) {
            const type = String(row[colIndex.projectType] || "").trim();
            if (type) {
                const stats = result.projectTypeStats.get(type) || { count: 0, cost: 0 };
                stats.count++;
                stats.cost += cost;
                result.projectTypeStats.set(type, stats);
            }
        }
    }
    result.hasData = (budgetFound || costFound) && (result.regionStats.size > 0 || result.statusStats.size > 0 || result.phaseStats.size > 0);
    return result;
}

// ========== TEK DASHBOARD SAYFASI ==========
async function createExecutiveDashboard(context, metrics, topProducts, channelPerformance, stockRisks, finance, warnings, constructionData) {
    let sheet = getOrCreateSheet(context, "EXECUTIVE_DASHBOARD");
    await context.sync();

    // Sayfayı temizle
    const used = sheet.getUsedRange();
    if (used) used.clear();

    // Başlık
    sheet.getRange("A1").values = [["📊 EXECUTIVE DASHBOARD – AI Özet Analiz"]];
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 16;

    let row = 3;

    // KPI'lar
    const kpis = [
        ["Toplam Adet", metrics.totalQuantity],
        ["Toplam Ciro (TL)", metrics.totalRevenue],
        ["Ortalama Adet", metrics.avgQuantity.toFixed(2)],
        ["Aykırı Değer Sayısı", metrics.outliers]
    ];
    const kpiRange = sheet.getRangeByIndexes(row, 0, kpis.length, 2);
    kpiRange.values = kpis;
    kpiRange.format.font.bold = true;
    row += kpis.length + 2;

    // Uyarılar
    if (warnings.length) {
        sheet.getRange(row, 0).values = [["⚠️ Uyarılar / Eksikler"]];
        sheet.getRange(row, 0).format.font.bold = true;
        row++;
        for (const w of warnings) {
            sheet.getRange(row, 0).values = [[w]];
            row++;
        }
        row++;
    }

    // Top ürünler (grafik önerisiyle)
    if (topProducts.length) {
        sheet.getRange(row, 0).values = [["🏆 En Çok Satan Ürünler (Grafik Önerisi: Sütun Grafik)"]];
        sheet.getRange(row, 0).format.font.bold = true;
        row++;
        const topRange = sheet.getRangeByIndexes(row, 0, topProducts.length, 2);
        topRange.values = topProducts.map(p => [p.product, p.quantity]);
        // Grafik ekle
        const chart = sheet.charts.add("columnClustered", topRange, "auto");
        chart.title.text = "Top Selling Products";
        chart.legend.position = "bottom";
        row += topProducts.length + 2;
    }

    // Kanal performansı
    if (channelPerformance.length) {
        sheet.getRange(row, 0).values = [["📢 Kanal / Bayi Performansı (Grafik Önerisi: Pasta veya Sütun)"]];
        sheet.getRange(row, 0).format.font.bold = true;
        row++;
        const channelRange = sheet.getRangeByIndexes(row, 0, channelPerformance.length, 2);
        channelRange.values = channelPerformance.map(c => [c.channel, c.quantity]);
        const chart = sheet.charts.add("pie", channelRange, "auto");
        chart.title.text = "Channel Distribution";
        chart.legend.position = "right";
        row += channelPerformance.length + 2;
    }

    // Stok riski
    if (stockRisks.length) {
        sheet.getRange(row, 0).values = [["⚠️ Stok Riskli Ürünler (Kritik / Düşük)"]];
        sheet.getRange(row, 0).format.font.bold = true;
        row++;
        sheet.getRangeByIndexes(row, 0, stockRisks.length, 3).values = stockRisks.map(r => [r.product, r.stock, r.risk]);
        row += stockRisks.length + 2;
    }

    // Finans
    if (finance) {
        sheet.getRange(row, 0).values = [["💰 Bütçe vs Gerçekleşen"]];
        sheet.getRange(row, 0).format.font.bold = true;
        row++;
        sheet.getRange(row, 0).values = [["Toplam Bütçe", finance.totalBudget]];
        sheet.getRange(row+1, 0).values = [["Toplam Gerçekleşen", finance.totalActual]];
        sheet.getRange(row+2, 0).values = [["Varyans", finance.variance]];
        if (finance.variance > 0) sheet.getRange(row+2, 1).format.font.color = "green";
        else if (finance.variance < 0) sheet.getRange(row+2, 1).format.font.color = "red";
        row += 5;
    }

    // İnşaat projesi verisi varsa
    if (constructionData.hasData) {
        sheet.getRange(row, 0).values = [["🏗️ İnşaat Projesi Analizi"]];
        sheet.getRange(row, 0).format.font.bold = true;
        row++;

        if (constructionData.regionStats.size) {
            sheet.getRange(row, 0).values = [["Bölgelere Göre Proje Sayısı (Grafik Önerisi: Pasta)"]];
            row++;
            const regionData = Array.from(constructionData.regionStats.entries()).map(([r, s]) => [r, s.count]);
            const regionRange = sheet.getRangeByIndexes(row, 0, regionData.length, 2);
            regionRange.values = regionData;
            const chart = sheet.charts.add("pie", regionRange, "auto");
            chart.title.text = "Projects by Region";
            row += regionData.length + 2;
        }

        if (constructionData.statusStats.size) {
            sheet.getRange(row, 0).values = [["Proje Durumları (Grafik Önerisi: Sütun)"]];
            row++;
            const statusData = Array.from(constructionData.statusStats.entries()).map(([s, st]) => [s, st.count]);
            const statusRange = sheet.getRangeByIndexes(row, 0, statusData.length, 2);
            statusRange.values = statusData;
            const chart = sheet.charts.add("columnClustered", statusRange, "auto");
            chart.title.text = "Project Status";
            row += statusData.length + 2;
        }

        if (constructionData.phaseStats.size) {
            sheet.getRange(row, 0).values = [["Proje Fazları (Grafik Önerisi: Çubuk)"]];
            row++;
            const phaseData = Array.from(constructionData.phaseStats.entries()).map(([p, st]) => [p, st.count]);
            const phaseRange = sheet.getRangeByIndexes(row, 0, phaseData.length, 2);
            phaseRange.values = phaseData;
            const chart = sheet.charts.add("barClustered", phaseRange, "auto");
            chart.title.text = "Project Phase";
            row += phaseData.length + 2;
        }
    }

    // Son not
    sheet.getRange(row, 0).values = [["💡 Not: Grafikler otomatik oluşturulmuştur. Dilerseniz üzerlerine tıklayarak düzenleyebilirsiniz."]];
    sheet.getRange(row, 0).format.font.italic = true;
    sheet.getRange("A:J").format.autofitColumns();
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
