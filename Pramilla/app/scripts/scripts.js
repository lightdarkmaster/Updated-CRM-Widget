/*******************************
 * CONSTANTS & HELPERS
 *******************************/
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const formatMonth = (date) => `${MONTH_NAMES[date.getMonth()]} ${date.getFullYear()}`;
const getAllMonthsOfYear = (year) => MONTH_NAMES.map(m => `${m} ${year}`);

const LEAD_SOURCES = [
    "Zoho Leads", "Zoho Partner", "Zoho CRM", "Zoho Partners 2024",
    "Zoho - Sutha", "Zoho - Hemanth", "Zoho - Sen", "Zoho - Audrey",
    "Zoho - Jacklyn", "Zoho - Adrian", "Zoho Partner Website", "Zoho - Chaitanya"
];

const ZOHO_SERVICES = ["Desk", "Workplace", "Projects", "Mail"];

// Get total weeks in a month (4 or 5)
function getWeeksInMonth(year, monthIndex) {
    const firstDay = new Date(year, monthIndex, 1).getDay();
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
    return Math.ceil((firstDay + daysInMonth) / 7);
}

// Get week number of the month for a date (1–5)
function getWeekOfMonth(date) {
    const firstDay = new Date(date.getFullYear(), date.getMonth(), 1).getDay();
    return Math.ceil((date.getDate() + firstDay) / 7);
}

/*************************************
 * OPTIMIZED COQL FETCH WITH PAGINATION
 *************************************/
async function fetchFilteredLeadsCOQL(year) {
    const startDate = `${year}-01-01`;
    const endDate = `${year}-12-31`;

    let allLeads = [];
    let offset = 0;
    const limit = 200;
    const maxOffset = 2000; // COQL limit

    console.log(`Starting COQL fetch for year ${year}...`);

    // Build OR conditions (COQL doesn't support IN with arrays)
    const sourceConditions = LEAD_SOURCES
        .map(src => `Lead_Source = '${src.replace(/'/g, "\\'")}'`)
        .join(" OR ");

    const serviceConditions = ZOHO_SERVICES
        .map(svc => `Zoho_Service = '${svc}'`)
        .join(" OR ");

    while (offset < maxOffset) {
        try {
            const query = `SELECT Created_Time, Lead_Source, Zoho_Service FROM Leads WHERE (${sourceConditions}) AND (${serviceConditions}) AND Created_Time >= '${startDate}' AND Created_Time <= '${endDate}' LIMIT ${limit} OFFSET ${offset}`;

            console.log(`Fetching offset ${offset}...`);

            const resp = await ZOHO.CRM.API.coql({ select_query: query });

            if (!resp?.data?.length) {
                console.log(`No more data at offset ${offset}`);
                break;
            }

            allLeads.push(...resp.data);
            console.log(`Fetched ${resp.data.length} leads (total: ${allLeads.length})`);

            // If we got less than limit, we've reached the end
            if (resp.data.length < limit) break;

            offset += limit;

        } catch (err) {
            console.error(`COQL error at offset ${offset}:`, err);
            
            // If COQL fails, try fallback method
            if (offset === 0) {
                console.warn("COQL failed, attempting fallback...");
                return await fetchLeadsFallback(year);
            }
            
            break;
        }
    }

    console.log(`Total leads fetched: ${allLeads.length}`);
    return allLeads;
}

/*************************************
 * FALLBACK: STANDARD API METHOD
 *************************************/
async function fetchLeadsFallback(year) {
    console.log("Using Standard API fallback method...");
    
    let allLeads = [];
    let page = 1;
    const perPage = 200;

    while (page <= 10) { // Limit to 10 pages for safety
        try {
            const config = {
                Entity: "Leads",
                sort_order: "asc",
                sort_by: "Created_Time",
                per_page: perPage,
                page: page
            };

            console.log(`Fetching page ${page}...`);

            const resp = await ZOHO.CRM.API.getAllRecords(config);

            if (!resp?.data?.length) break;

            // Filter by year and criteria
            const filtered = resp.data.filter(lead => {
                if (!lead.Created_Time) return false;

                try {
                    const date = new Date(lead.Created_Time.split(/[+Z]/)[0].replace("T", " "));
                    if (date.getFullYear() !== year) return false;

                    return LEAD_SOURCES.includes(lead.Lead_Source) && 
                           ZOHO_SERVICES.includes(lead.Zoho_Service);
                } catch (e) {
                    return false;
                }
            });

            allLeads.push(...filtered);

            if (!resp.info?.more_records) break;

            page++;

        } catch (err) {
            console.error(`API error at page ${page}:`, err);
            break;
        }
    }

    console.log(`Fallback method fetched ${allLeads.length} leads`);
    return allLeads;
}

/********************************************
 * GROUP LEADS BY MONTH & WEEK (4–5 WEEKS)
 ********************************************/
function groupLeadsByMonthWeek(leads, year) {
    const grouped = {};
    
    MONTH_NAMES.forEach((m, idx) => {
        const monthKey = `${m} ${year}`;
        const weeks = getWeeksInMonth(year, idx);
        grouped[monthKey] = {};
        for (let w = 1; w <= weeks; w++) grouped[monthKey][w] = 0;
    });

    leads.forEach(lead => {
        const createdTime = lead.Created_Time;
        if (!createdTime) return;

        try {
            const date = new Date(createdTime.split(/[+Z]/)[0].replace("T", " "));
            if (isNaN(date.getTime()) || date.getFullYear() !== year) return;

            const monthKey = formatMonth(date);
            const week = getWeekOfMonth(date);
            if (grouped[monthKey]?.[week] !== undefined) {
                grouped[monthKey][week]++;
            }
        } catch (e) {
            console.error("Date parsing error:", e, createdTime);
        }
    });

    return grouped;
}

/*************************************
 * PERCENT CHANGE HELPER
 *************************************/
function getPercentChange(current, previous) {
    if (previous == null) return "";
    if (previous === 0) {
        return current > 0
            ? ` <span style="color:green;font-weight:bold;">(+∞)</span>`
            : ` <span style="color:gray;">(—)</span>`;
    }
    const change = (((current - previous) / previous) * 100).toFixed(1);
    const color = change > 0 ? "green" : change < 0 ? "red" : "gray";
    return ` <span style="color:${color};font-weight:bold;">(${change > 0 ? "+" : ""}${change}%)</span>`;
}

/*************************************
 * TABLE RENDERING (DYNAMIC WEEKS)
 *************************************/
function renderTable(monthlyWeeklyCounts, year, totalFiltered) {
    const tableContainer = document.querySelector("#tableContainer");
    tableContainer.innerHTML = `
        <div style="margin:20px;font-family:Arial,sans-serif;">
            <h2 style="color:#333;">Lead Generation Report - ${year}</h2>
            <table id="leadsTable" style="border-collapse:collapse;width:100%;margin:20px 0;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
                <thead></thead>
                <tbody></tbody>
            </table>
            <div id="footerNote"></div>
        </div>`;

    const table = document.querySelector("#leadsTable");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");

    const months = getAllMonthsOfYear(year);
    const maxWeeks = Math.max(...months.map(m => Object.keys(monthlyWeeklyCounts[m]).length));

    // Table Header
    thead.innerHTML = `<tr style="background:#4CAF50;color:white;">
        <th style="padding:12px;border:1px solid #ddd;text-align:center;width:100px;">Week</th>
        ${months.map(m => `<th style="padding:12px;border:1px solid #ddd;text-align:center;">${m}</th>`).join("")}
    </tr>`;

    // Week Rows
    for (let week = 1; week <= maxWeeks; week++) {
        const row = months.map((m, idx) => {
            const count = monthlyWeeklyCounts[m]?.[week];
            
            if (count === undefined) {
                return `<td style="padding:10px;border:1px solid #ddd;text-align:center;background:#f5f5f5;"></td>`;
            }

            let prev = null;
            if (week > 1) {
                prev = monthlyWeeklyCounts[m]?.[week - 1];
            } else if (week === 1 && idx > 0) {
                const prevMonth = months[idx - 1];
                const prevWeeks = Object.keys(monthlyWeeklyCounts[prevMonth]);
                prev = monthlyWeeklyCounts[prevMonth]?.[Math.max(...prevWeeks)];
            }
            
            return `<td style="padding:10px;border:1px solid #ddd;text-align:center;">${count}${getPercentChange(count, prev)}</td>`;
        }).join("");

        tbody.innerHTML += `<tr style="background:${week % 2 ? "white" : "#f9f9f9"};">
            <td style="padding:10px;border:1px solid #ddd;font-weight:bold;text-align:center;">Week ${week}</td>
            ${row}
        </tr>`;
    }

    // Monthly Totals
    let grandTotal = 0;
    const totalRow = months.map((m, i) => {
        const total = Object.values(monthlyWeeklyCounts[m] || {}).reduce((a, b) => a + b, 0);
        grandTotal += total;
        const prev = i > 0 
            ? Object.values(monthlyWeeklyCounts[months[i - 1]] || {}).reduce((a, b) => a + b, 0) 
            : null;
        return `<td style="padding:12px;border:1px solid #ddd;text-align:center;background:#e8f5e9;"><strong>${total}${getPercentChange(total, prev)}</strong></td>`;
    }).join("");

    tbody.innerHTML += `<tr style="background:#c8e6c9;font-weight:bold;">
        <td style="padding:12px;border:1px solid #ddd;text-align:center;">Monthly Total</td>
        ${totalRow}
    </tr>`;

    document.querySelector("#footerNote").innerHTML = `
        <div style="background:#f5f5f5;padding:20px;border-radius:8px;margin-top:20px;border-left:4px solid #4CAF50;">
            <h4 style="margin:0 0 15px;color:#333;font-size:1em">Lead Generation Summary for ${year}</h4>
            <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:15px;font-size:0.9em">
                <div><strong>Total Leads:</strong> <span style="color:#2196F3;font-size:1.3em;font-weight:bold;">${grandTotal}</span></div>
                <div><strong>Total Fetched:</strong> ${totalFiltered}</div>
                <div><strong>Report Period:</strong> Jan - Dec ${year}</div>
                <div><strong>Generated:</strong> ${new Date().toLocaleString()}</div>
            </div>
            <p style="margin-top:15px;color:#666;font-size:0.85em;border-top:1px solid #ddd;padding-top:10px;">
                Week-over-week and month-over-month percentage changes shown in <span style="color:green;">green</span> (increase) or <span style="color:red;">red</span> (decrease).
            </p>
        </div>`;
}

/*************************************
 * PAGE LOAD
 *************************************/
ZOHO.embeddedApp.on("PageLoad", async () => {
    const year = new Date().getFullYear();

    document.body.innerHTML = `
        <div id="loadingDiv" style="text-align:center;padding:40px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;border-radius:12px;margin:20px;box-shadow:0 4px 12px rgba(0,0,0,0.15);">
            <h2 style="margin:0 0 10px;">Loading Lead Data for ${year}</h2>
            <p style="margin:10px 0;">Applying Lead Source and Zoho Service filters...</p>
            <div style="margin-top:20px;">
                <div style="display:inline-block;width:40px;height:40px;border:4px solid rgba(255,255,255,0.3);border-top:4px solid white;border-radius:50%;animation:spin 1s linear infinite;"></div>
            </div>
            <style>
                @keyframes spin { 
                    0% { transform: rotate(0deg); } 
                    100% { transform: rotate(360deg); } 
                }
            </style>
        </div>
        <div id="tableContainer"></div>`;

    try {
        const leads = await fetchFilteredLeadsCOQL(year);

        if (!leads || leads.length === 0) {
            document.body.innerHTML = `
                <div style="text-align:center;padding:40px;background:#fff3cd;border-radius:12px;margin:20px;border-left:5px solid #ffc107;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
                    <h2 style="color:#856404;margin:0 0 10px;">⚠️ No Matching Leads Found</h2>
                    <p style="color:#856404;margin:10px 0;">No leads found matching the specified criteria for ${year}.</p>
                    <div style="margin-top:20px;padding:15px;background:white;border-radius:8px;text-align:left;">
                        <strong>Filter Criteria:</strong><br>
                        <strong>Lead Sources:</strong> ${LEAD_SOURCES.join(", ")}<br>
                        <strong>Zoho Services:</strong> ${ZOHO_SERVICES.join(", ")}
                    </div>
                </div>`;
            return;
        }

        const grouped = groupLeadsByMonthWeek(leads, year);
        renderTable(grouped, year, leads.length);

        document.getElementById("loadingDiv").remove();

    } catch (err) {
        console.error("Critical error:", err);
        document.body.innerHTML = `
            <div style="text-align:center;padding:40px;background:#f8d7da;border-radius:12px;margin:20px;border-left:5px solid #dc3545;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
                <h2 style="color:#721c24;margin:0 0 10px;">Error Loading Report</h2>
                <p style="color:#721c24;margin:10px 0;"><strong>${err.message}</strong></p>
                <button onclick="location.reload()" style="margin-top:15px;padding:12px 24px;background:#007bff;color:white;border:none;border-radius:6px;cursor:pointer;font-size:1em;box-shadow:0 2px 4px rgba(0,0,0,0.2);transition:all 0.3s;">
                    Retry
                </button>
            </div>`;
    }
});

ZOHO.embeddedApp.init();