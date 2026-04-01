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

            // Filtre değerlerini al (tarih yok)
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
            const warnings = collectWarnings(headers, colIndex);
            const topProducts = getTopProducts(filteredRows, colIndex);
            const channelPerformance = getChannelPerformance(filteredRows, colIndex);
            const stockRisks = getStockRisk(filteredRows, colIndex);
            const financeSummary = getFinanceSummary(filteredRows, colIndex);
            const constructionData = prepareConstructionData(filteredRows, colIndex);

            // Grafik önerileri
            const chartSuggestions = getChartSuggestions(topProducts, channelPerformance, constructionData);

            // Task pane özet metnini oluştur
            let resultText = `✅ ANALİZ TAMAMLANDI!\n\n`;
            resultText += `📄 Sayfa: ${sheet.name}\n`;
            resultText += `📍 Aralık: ${usedRange.address}\n`;
            resultText += `📊 Satır: ${usedRange.rowCount} | Sütun: ${usedRange.columnCount}\n\n`;
            resultText += `🔍 Başlıklar: ${headers.join(", ")}\n\n`;
            resultText += `📈 GENEL METRİKLER\n`;
            resultText += `   • Toplam Adet: ${metrics.totalQuantity}\n`;
            resultText += `   • Toplam Ciro: ${metrics.totalRevenue.toLocaleString()} TL\n`;
            resultText += `   • Ortalama Adet: ${metrics.avgQuantity.toFixed(2)}\n`;
            resultText += `   • Aykırı Değer Sayısı: ${metrics.outliers}\n\n`;

            if (warnings.length) {
                resultText += `⚠️ UYARILAR / EKSİKLER\n`;
                warnings.forEach(w => { resultText += `   • ${w}\n`; });
                resultText += `\n`;
            }

            if (topProducts.length) {
                resultText += `🏆 EN ÇOK SATAN ÜRÜNLER\n`;
                topProducts.forEach(p => { resultText += `   • ${p.product}: ${p.quantity} adet\n`; });
                resultText += `\n`;
            }

            if (channelPerformance.length) {
                resultText += `📢 KANAL / BAYİ PERFORMANSI\n`;
                channelPerformance.forEach(c => { resultText += `   • ${c.channel}: ${c.quantity} adet\n`; });
                resultText += `\n`;
            }

            if (stockRisks.length) {
                resultText += `⚠️ STOK RİSKLİ ÜRÜNLER\n`;
                stockRisks.forEach(r => { resultText += `   • ${r.product}: ${r.stock} adet (${r.risk})\n`; });
                resultText += `\n`;
            }

            if (financeSummary) {
                resultText += `💰 BÜTÇE ANALİZİ\n`;
                resultText += `   • Toplam Bütçe: ${financeSummary.totalBudget.toLocaleString()} TL\n`;
                resultText += `   • Toplam Gerçekleşen: ${financeSummary.totalActual.toLocaleString()} TL\n`;
                resultText += `   • Varyans: ${financeSummary.variance.toLocaleString()} TL\n`;
                resultText += `\n`;
            }

            if (constructionData.hasData) {
                resultText += `🏗️ İNŞAAT PROJE ANALİZİ\n`;
                if (constructionData.regionStats.size) {
                    resultText += `   Bölge Dağılımı:\n`;
                    constructionData.regionStats.forEach((stats, region) => {
                        resultText += `      • ${region}: ${stats.count} proje\n`;
                    });
                }
                if (constructionData.statusStats.size) {
                    resultText += `   Durum Dağılımı:\n`;
                    constructionData.statusStats.forEach((stats, status) => {
                        resultText += `      • ${status}: ${stats.count} proje\n`;
                    });
                }
                if (constructionData.phaseStats.size) {
                    resultText += `   Faz Dağılımı:\n`;
                    constructionData.phaseStats.forEach((stats, phase) => {
                        resultText += `      • ${phase}: ${stats.count} proje\n`;
                    });
                }
                resultText += `\n`;
            }

            if (chartSuggestions.length) {
                resultText += `📊 GRAFİK ÖNERİLERİ\n`;
                chartSuggestions.forEach(s => { resultText += `   • ${s}\n`; });
                resultText += `\n`;
            }

            resultText += `💡 Not: Tüm analizler bu panelde gösterilmektedir. Excel’e herhangi bir sayfa eklenmemiştir.`;

            showResult(resultText);

            // Filtre dropdown'larını güncelle (kullanıcı sonra seçim yapabilir)
            updateFilterDropdowns(headers, dataRows, colIndex);

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

function collectWarnings(headers, colIndex) {
    const warnings = [];
    if (colIndex.quantity === -1) warnings.push("⚠️ 'Adet' sütunu bulunamadı. Satış adedi analizi yapılamıyor.");
    if (colIndex.revenue === -1) warnings.push("⚠️ 'Ciro/Tutar' sütunu bulunamadı. Gelir analizi yapılamıyor.");
    if (colIndex.channel === -1) warnings.push("⚠️ 'Kanal/Bayi' sütunu bulunamadı. Kanal performansı gösterilemiyor.");
    if (colIndex.product === -1) warnings.push("⚠️ 'Ürün' sütunu bulunamadı. Ürün bazlı analiz yapılamıyor.");
    if (colIndex.stock === -1) warnings.push("ℹ️ 'Stok' sütunu bulunamadı. Stok riski analizi atlandı.");
    if (colIndex.budget === -1 || colIndex.actual === -1) warnings.push("ℹ️ Bütçe/Gerçekleşen sütunları eksik. Finans analizi sınırlı.");
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

function getChartSuggestions(topProducts, channelPerformance, constructionData) {
    const suggestions = [];
    if (topProducts.length) {
        suggestions.push("Top ürünler için sütun grafik kullanarak satış adetlerini karşılaştırabilirsiniz.");
    }
    if (channelPerformance.length) {
        suggestions.push("Kanal/bayi dağılımını göstermek için pasta grafik etkili olacaktır.");
    }
    if (constructionData.hasData) {
        if (constructionData.regionStats.size) {
            suggestions.push("Projelerin bölgelere göre dağılımını pasta grafikle görselleştirebilirsiniz.");
        }
        if (constructionData.statusStats.size) {
            suggestions.push("Proje durumlarını sütun grafikle izleyebilirsiniz.");
        }
        if (constructionData.phaseStats.size) {
            suggestions.push("Proje fazlarını çubuk grafikle karşılaştırabilirsiniz.");
        }
    }
    if (suggestions.length === 0) {
        suggestions.push("Grafik önerisi için yeterli veri bulunamadı (ürün, kanal veya inşaat verisi yok).");
    }
    return suggestions;
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
        analyzeBtn.textContent = show ? "⏳ Analiz Ediliyor..." : "📊 Veriyi Analiz Et";
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
