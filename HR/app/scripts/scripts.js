const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const formatMonth = (date) => `${MONTH_NAMES[date.getMonth()]} ${date.getFullYear()}`;
const getAllMonthsOfYear = (year) => MONTH_NAMES.map(m => `${m} ${year}`);

// Helper function to get the number of weeks in a month
function getWeeksInMonth(year, month) {
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const lastDate = lastDay.getDate();
    
    // Calculate weeks based on the last date
    return Math.ceil(lastDate / 7);
}

async function fetchFilteredLeads() {
    const leadSources = [
        "Zoho Leads", "Zoho Partner", "Zoho CRM", "Zoho Partners 2024",
        "Zoho - Sutha", "Zoho - Hemanth", "Zoho - Sen", "Zoho - Audrey",
        "Zoho - Jacklyn", "Zoho - Adrian", "Zoho Partner Website", "Zoho - Chaitanya"
    ];
    const zohoServices = ["People", "Recruit", "Expense"];

    let allData = [];
    let page = 1;
    let hasMore = true;

    while (hasMore) {
        try {
            const resp = await ZOHO.CRM.API.getAllRecords({
                Entity: "Leads",
                per_page: 200,
                page: page
            });

            if (resp && resp.data && resp.data.length > 0) {
                const filteredData = resp.data.filter(lead => {
                    const leadSource = lead.Lead_Source || lead['Lead Source'] || lead.lead_source || "";
                    const zohoService = lead.Zoho_Service || lead['Zoho Service'] || lead.zoho_service || "";
                    
                    const matchesLeadSource = leadSources.includes(leadSource);
                    const matchesZohoService = zohoServices.includes(zohoService);
                    
                    return matchesLeadSource && matchesZohoService;
                });

                allData.push(...filteredData);
                hasMore = resp.data.length === 200;
                page++;
                
                if (page > 100) break;
            } else {
                hasMore = false;
            }
        } catch (err) {
            console.error(`Error fetching page ${page}:`, err);
            break;
        }
    }

    return allData;
}

function groupLeadsByMonthWeek(leads, year) {
    // Initialize with dynamic week counts per month
    const grouped = {};
    
    for (let monthIndex = 0; monthIndex < 12; monthIndex++) {
        const monthKey = `${MONTH_NAMES[monthIndex]} ${year}`;
        const weeksInMonth = getWeeksInMonth(year, monthIndex);
        
        grouped[monthKey] = {
            weeks: {},
            maxWeeks: weeksInMonth
        };
        
        // Initialize all weeks for this month
        for (let w = 1; w <= weeksInMonth; w++) {
            grouped[monthKey].weeks[w] = 0;
        }
    }

    leads.forEach(lead => {
        const createdTimeValue = lead.Created_Time || lead.created_time || 
                                lead.Created_Date || lead.created_date;
        
        if (!createdTimeValue) return;

        try {
            let dateString = createdTimeValue;
            if (typeof dateString === 'string') {
                dateString = dateString.split(/[+Z]/)[0].replace('T', ' ');
            }
            
            const createdDate = new Date(dateString);
            
            if (isNaN(createdDate.getTime()) || createdDate.getFullYear() !== year) return;

            const monthKey = formatMonth(createdDate);
            if (!grouped[monthKey]) return;

            const week = Math.ceil(createdDate.getDate() / 7);
            const maxWeeks = grouped[monthKey].maxWeeks;
            
            // Ensure week doesn't exceed the actual weeks in the month
            const adjustedWeek = Math.min(week, maxWeeks);
            
            if (grouped[monthKey].weeks[adjustedWeek] !== undefined) {
                grouped[monthKey].weeks[adjustedWeek]++;
            }
        } catch (dateErr) {
            console.error('Date parsing error:', dateErr);
        }
    });

    return grouped;
}

function getPercentChange(current, previous) {
    if (previous == null) return "";
    if (previous === 0) {
        if (current > 0) return ` <span style="color:green; font-weight:bold;">(+∞)</span>`;
        return ` <span style="color:gray;">(0%)</span>`;
    }
    const change = (((current - previous) / previous) * 100).toFixed(1);
    const color = change > 0 ? "green" : change < 0 ? "red" : "gray";
    return ` <span style="color:${color}; font-weight:bold;">(${change > 0 ? "+" : ""}${change}%)</span>`;
}

function renderTable(monthlyWeeklyCounts, year, totalFiltered) {
    let table = document.querySelector("#leadsTable");
    if (!table) {
        document.body.innerHTML += `
            <div style="margin: 20px;">
                <h2>Lead Generation Report - ${year}</h2>
                <table id="leadsTable" style="border-collapse: collapse; width: 100%; margin: 20px 0;">
                    <thead></thead><tbody></tbody>
                </table>
                <div id="footerNote"></div>
            </div>`;
        table = document.querySelector("#leadsTable");
    }

    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");
    thead.innerHTML = tbody.innerHTML = "";

    const months = getAllMonthsOfYear(year);
    
    // Find the maximum number of weeks across all months
    const maxWeeks = Math.max(...months.map(m => monthlyWeeklyCounts[m]?.maxWeeks || 4));

    thead.innerHTML = `<tr style="background:#4CAF50;color:white;">
        <th style="padding:12px;border:1px solid #ddd;text-align:center;">Week</th>
        ${months.map(m => `<th style="padding:12px;border:1px solid #ddd;text-align:center;">${m}</th>`).join("")}
    </tr>`;

    for (let week = 1; week <= maxWeeks; week++) {
        const row = months.map((m, monthIndex) => {
            const monthData = monthlyWeeklyCounts[m];
            const hasThisWeek = monthData && week <= monthData.maxWeeks;
            const count = hasThisWeek ? (monthData.weeks[week] || 0) : null;
            
            if (count === null) {
                return `<td style="padding:10px;border:1px solid #ddd;text-align:center;background:#f5f5f5;">—</td>`;
            }
            
            let prev = null;
            
            if (week > 1) {
                prev = monthData.weeks[week - 1] || 0;
            } else if (week === 1 && monthIndex > 0) {
                const prevMonth = months[monthIndex - 1];
                const prevMonthData = monthlyWeeklyCounts[prevMonth];
                if (prevMonthData) {
                    prev = prevMonthData.weeks[prevMonthData.maxWeeks] || 0;
                }
            }
            
            return `<td style="padding:10px;border:1px solid #ddd;text-align:center;">${count}${getPercentChange(count, prev)}</td>`;
        }).join("");

        tbody.innerHTML += `<tr style="background:${week % 2 ? "white" : "#f9f9f9"};">
            <td style="padding:10px;border:1px solid #ddd;font-weight:bold;text-align:center;">Week ${week}</td>
            ${row}
        </tr>`;
    }

    let grandTotal = 0;
    const totalRow = months.map((m, i) => {
        const monthData = monthlyWeeklyCounts[m];
        const total = monthData ? Object.values(monthData.weeks).reduce((a, b) => a + b, 0) : 0;
        grandTotal += total;
        
        const prevMonthData = i > 0 ? monthlyWeeklyCounts[months[i - 1]] : null;
        const prev = prevMonthData ? Object.values(prevMonthData.weeks).reduce((a, b) => a + b, 0) : null;
        
        return `<td style="padding:12px;border:1px solid #ddd;text-align:center;"><strong>${total}${getPercentChange(total, prev)}</strong></td>`;
    }).join("");

    tbody.innerHTML += `<tr style="background:#f0f1f2;color:white;font-weight:bold;">
        <td style="padding:12px;border:1px solid #ddd;text-align:center;">Monthly Total</td>${totalRow}
    </tr>`;

    document.querySelector("#footerNote").innerHTML = `
        <div style="background:#f5f5f5;padding:20px;border-radius:8px;margin-top:20px;">
            <h4 style="margin:0 0 15px;color:#333;font-size:0.9em">Lead Generation Summary for ${year}</h4>
            <div style="display:flex;flex-wrap:wrap;gap:20px;font-size:0.9em">
                <div><strong>Filtered Leads:</strong> <span style="color:#2196F3;font-size:1.2em;">${grandTotal}</span></div>
                <div><strong>Total Fetched:</strong> ${totalFiltered}</div>
                <div><strong>Report Period:</strong> Jan - Dec ${year}</div>
                <div><strong>Generated:</strong> ${new Date().toLocaleString()}</div>
            </div>
            <p style="margin-top:15px;color:#666;font-size:0.9em;">
                Week-over-week and month-over-month percentage changes shown. Months with 5 weeks are fully supported.
            </p>
        </div>`;
}

ZOHO.embeddedApp.on("PageLoad", async () => {
    const targetYear = new Date().getFullYear();
    document.body.innerHTML = `
        <div id="loadingDiv" style="text-align:center;padding:40px;background:#e3f2fd;border-radius:8px;margin:20px;">
            <h2>Loading Filtered Lead Data for ${targetYear}</h2>
            <p>Applying Lead Source and Zoho Service filters...</p>
        </div>
        <table id="leadsTable" style="border-collapse:collapse;width:100%;margin:20px 0;"><thead></thead><tbody></tbody></table>
        <div id="footerNote"></div>`;

    try {
        const filteredLeads = await fetchFilteredLeads();
        
        if (!filteredLeads.length) {
            document.body.innerHTML = `
                <div style="text-align:center;padding:40px;background:#fff3cd;border-radius:8px;margin:20px;border-left:5px solid #ffc107;">
                    <h2>No Matching Leads Found</h2>
                    <p>No leads found matching the specified Lead Source and Zoho Service criteria.</p>
                </div>`;
            return;
        }

        const monthlyWeeklyCounts = groupLeadsByMonthWeek(filteredLeads, targetYear);
        renderTable(monthlyWeeklyCounts, targetYear, filteredLeads.length);
        
        document.getElementById("loadingDiv").style.display = "none";

    } catch (err) {
        document.body.innerHTML = `
            <div style="text-align:center;padding:40px;background:#f8d7da;border-radius:8px;margin:20px;border-left:5px solid #dc3545;">
                <h2>Error Loading Report</h2>
                <p><strong>${err.message}</strong></p>
                <button onclick="location.reload()" style="padding:10px 20px;background:#007bff;color:white;border:none;border-radius:5px;cursor:pointer;">
                    Retry
                </button>
            </div>`;
    }
});

ZOHO.embeddedApp.init();