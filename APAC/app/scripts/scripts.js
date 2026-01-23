/************************************************
 * CONFIG
 ************************************************/
const ZOHO_SERVICE_API = "Zoho_Service__c"; // 

const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const LEAD_SOURCES = [
    "Zoho Leads", "Zoho Partner", "Zoho CRM", "Zoho Partners 2024",
    "Zoho - Sutha", "Zoho - Hemanth", "Zoho - Sen", "Zoho - Audrey",
    "Zoho - Jacklyn", "Zoho - Adrian", "Zoho Partner Website", "Zoho - Chaitanya"
];

const ZOHO_SERVICES = ["CRM", "CRMPlus", "One", "Bigin"];

/************************************************
 * DATE HELPERS
 ************************************************/
const formatMonth = date =>
    `${MONTH_NAMES[date.getMonth()]} ${date.getFullYear()}`;

const getAllMonthsOfYear = year =>
    MONTH_NAMES.map(m => `${m} ${year}`);

/************************************************
 * FETCH USING STANDARD API (Most Reliable)
 ************************************************/
async function fetchFilteredLeadsAPI(targetYear) {
    const startDate = `${targetYear}-01-01`;
    const endDate = `${targetYear}-12-31`;

    let allData = [];
    let page = 1;
    const perPage = 200;
    let hasMore = true;

    console.log(`Fetching leads for year ${targetYear}...`);

    while (hasMore) {
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

            console.log(`Page ${page} response:`, resp);

            if (!resp?.data || resp.data.length === 0) {
                hasMore = false;
                break;
            }

            // Filter by date and criteria
            const filtered = resp.data.filter(lead => {
                if (!lead.Created_Time) return false;

                try {
                    const dateStr = lead.Created_Time.split(/[+Z]/)[0].replace("T", " ");
                    const createdDate = new Date(dateStr);
                    const createdYear = createdDate.getFullYear();

                    if (createdYear !== targetYear) return false;

                    const validSource = LEAD_SOURCES.includes(lead.Lead_Source);
                    const validService = ZOHO_SERVICES.includes(lead[ZOHO_SERVICE_API]);

                    return validSource && validService;
                } catch (e) {
                    console.error("Error filtering lead:", e, lead);
                    return false;
                }
            });

            allData.push(...filtered);

            console.log(`Page ${page}: ${resp.data.length} total, ${filtered.length} matched criteria`);

            hasMore = resp.info?.more_records === true;
            page++;

            // Safety limit
            if (page > 50) {
                console.warn("Reached page limit of 50");
                break;
            }

        } catch (err) {
            console.error(`Error fetching page ${page}:`, err);
            hasMore = false;
        }
    }

    console.log(`Total filtered leads: ${allData.length}`);
    return allData;
}

/************************************************
 * ALTERNATIVE: Simple COQL
 ************************************************/
async function fetchFilteredLeadsCOQL(targetYear) {
    let allData = [];
    let hasMore = true;
    let page = 1;

    console.log("Attempting COQL method...");

    while (hasMore && page <= 10) {
        try {
            const offset = (page - 1) * 200;
            
            const selectQuery = `SELECT id, Created_Time, Lead_Source, ${ZOHO_SERVICE_API} FROM Leads WHERE Created_Time >= '${targetYear}-01-01' AND Created_Time <= '${targetYear}-12-31' LIMIT 200 OFFSET ${offset}`;

            console.log("COQL Query:", selectQuery);

            const response = await ZOHO.CRM.API.coql({
                select_query: selectQuery
            });

            console.log("COQL Response:", response);

            if (!response?.data || response.data.length === 0) {
                hasMore = false;
                break;
            }

            // Filter in JavaScript
            const filtered = response.data.filter(lead => {
                const validSource = LEAD_SOURCES.includes(lead.Lead_Source);
                const validService = ZOHO_SERVICES.includes(lead[ZOHO_SERVICE_API]);
                return validSource && validService;
            });

            allData.push(...filtered);

            if (response.data.length < 200) {
                hasMore = false;
            }

            page++;

        } catch (err) {
            console.error("COQL error:", err);
            throw err;
        }
    }

    console.log(`COQL fetched ${allData.length} leads`);
    return allData;
}

/************************************************
 * GROUP LEADS BY MONTH & WEEK
 ************************************************/
function groupLeadsByMonthWeek(leads, year) {
    const grouped = Object.fromEntries(
        getAllMonthsOfYear(year).map(m => [m, {1:0, 2:0, 3:0, 4:0}])
    );

    console.log(`Grouping ${leads.length} leads...`);

    leads.forEach(lead => {
        const createdTime = lead.Created_Time;
        if (!createdTime) return;

        try {
            const dateStr = createdTime.split(/[+Z]/)[0].replace("T", " ");
            const date = new Date(dateStr);

            if (isNaN(date) || date.getFullYear() !== year) return;

            const monthKey = formatMonth(date);
            if (!grouped[monthKey]) return;

            const week = Math.min(4, Math.ceil(date.getDate() / 7));
            grouped[monthKey][week]++;

        } catch (e) {
            console.error("Date parse error:", e, "for", createdTime);
        }
    });

    console.log("Grouped data:", grouped);
    return grouped;
}

/************************************************
 * PERCENT CHANGE CALCULATION
 ************************************************/
function getPercentChange(current, previous) {
    if (previous == null) return "";

    if (previous === 0) {
        return current > 0
            ? ` <span style="color:green;font-weight:bold;">(+∞)</span>`
            : ` <span style="color:gray;">(—)</span>`;
    }

    const pct = (((current - previous) / previous) * 100).toFixed(1);
    const color = pct > 0 ? "green" : pct < 0 ? "red" : "gray";

    return ` <span style="color:${color};font-weight:bold;">(${pct > 0 ? "+" : ""}${pct}%)</span>`;
}

/************************************************
 * RENDER REPORT TABLE
 ************************************************/
function renderTable(monthlyWeeklyCounts, year, totalFiltered, totalFetched) {
    document.body.innerHTML = `
        <div style="margin:20px;font-family:Arial,sans-serif;">
            <h2>Lead Generation Report – ${year}</h2>
            <table id="leadsTable" style="border-collapse:collapse;width:100%;box-shadow:0 2px 4px rgba(0,0,0,0.1);">
                <thead></thead>
                <tbody></tbody>
            </table>
            <div id="footerNote"></div>
        </div>`;

    const table = document.querySelector("#leadsTable");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");

    const months = getAllMonthsOfYear(year);

    thead.innerHTML = `
        <tr style="background:#4CAF50;color:white;">
            <th style="padding:12px;border:1px solid #ddd;">Week</th>
            ${months.map(m => `<th style="padding:12px;border:1px solid #ddd;">${m}</th>`).join("")}
        </tr>`;

    for (let week = 1; week <= 4; week++) {
        const rowHtml = months.map((m, i) => {
            const count = monthlyWeeklyCounts[m][week];
            let prev = null;

            if (week > 1) {
                prev = monthlyWeeklyCounts[m][week - 1];
            } else if (i > 0) {
                prev = monthlyWeeklyCounts[months[i - 1]][4];
            }

            return `<td style="padding:10px;border:1px solid #ddd;text-align:center;">${count}${getPercentChange(count, prev)}</td>`;
        }).join("");

        tbody.innerHTML += `
            <tr style="background:${week % 2 ? "#fff" : "#f9f9f9"};">
                <td style="padding:10px;border:1px solid #ddd;font-weight:bold;">Week ${week}</td>
                ${rowHtml}
            </tr>`;
    }

    let grandTotal = 0;

    const totalRow = months.map((m, i) => {
        const total = Object.values(monthlyWeeklyCounts[m]).reduce((a, b) => a + b, 0);
        grandTotal += total;

        const prev = i > 0
            ? Object.values(monthlyWeeklyCounts[months[i - 1]]).reduce((a, b) => a + b, 0)
            : null;

        return `<td style="padding:12px;border:1px solid #ddd;text-align:center;"><strong>${total}${getPercentChange(total, prev)}</strong></td>`;
    }).join("");

    tbody.innerHTML += `
        <tr style="background:#e8f5e9;font-weight:bold;">
            <td style="padding:12px;border:1px solid #ddd;">Monthly Total</td>
            ${totalRow}
        </tr>`;

    document.querySelector("#footerNote").innerHTML = `
        <div style="background:#f5f5f5;padding:20px;border-radius:8px;margin-top:20px;">
            <strong>Filtered Leads:</strong> ${grandTotal}<br>
            <strong>Total Processed:</strong> ${totalFetched}<br>
            <strong>Report Period:</strong> Jan–Dec ${year}<br>
            <strong>Generated:</strong> ${new Date().toLocaleString()}
        </div>`;
}

/************************************************
 * PAGE LOAD
 ************************************************/
ZOHO.embeddedApp.on("PageLoad", async () => {
    const targetYear = new Date().getFullYear();

    document.body.innerHTML = `
        <div style="padding:40px;text-align:center;font-family:Arial,sans-serif;">
            <h2>Loading Lead Generation Report (${targetYear})</h2>
            <p>Fetching data from Zoho CRM…</p>
            <div style="margin-top:20px;">
                <div style="display:inline-block;width:50px;height:50px;border:5px solid #f3f3f3;border-top:5px solid #4CAF50;border-radius:50%;animation:spin 1s linear infinite;"></div>
            </div>
            <style>@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }</style>
        </div>`;

    try {
        let leads = [];
        let method = "Standard API";
        
        // Use Standard API as primary method (more reliable)
        try {
            console.log("Using Standard API method...");
            leads = await fetchFilteredLeadsAPI(targetYear);
        } catch (apiError) {
            console.error("Standard API failed:", apiError);
            
            // Try COQL as fallback
            try {
                console.log("Falling back to COQL...");
                method = "COQL";
                leads = await fetchFilteredLeadsCOQL(targetYear);
            } catch (coqlError) {
                console.error("COQL also failed:", coqlError);
                throw new Error("Both API methods failed. Check console for details.");
            }
        }

        console.log(`Successfully fetched ${leads.length} leads using ${method}`);

        if (!leads || leads.length === 0) {
            document.body.innerHTML = `
                <div style="padding:40px;text-align:center;font-family:Arial,sans-serif;">
                    <h2>No Matching Leads Found</h2>
                    <p>No leads match the specified criteria for ${targetYear}.</p>
                    <div style="margin-top:20px;padding:20px;background:#f5f5f5;border-radius:8px;text-align:left;">
                        <strong>Filter Criteria:</strong><br>
                        <strong>Lead Sources:</strong> ${LEAD_SOURCES.join(", ")}<br>
                        <strong>Zoho Services:</strong> ${ZOHO_SERVICES.join(", ")}<br>
                        <strong>Field Name:</strong> ${ZOHO_SERVICE_API}
                    </div>
                </div>`;
            return;
        }

        const grouped = groupLeadsByMonthWeek(leads, targetYear);
        renderTable(grouped, targetYear, leads.length, leads.length);

    } catch (err) {
        console.error("Critical error:", err);
        document.body.innerHTML = `
            <div style="padding:40px;text-align:center;color:red;font-family:Arial,sans-serif;">
                <h2>Error Loading Report</h2>
                <p><strong>${err.message}</strong></p>
                <div style="margin-top:20px;padding:20px;background:#fff3cd;border:1px solid #ffc107;border-radius:8px;text-align:left;color:#856404;">
                    <strong>Troubleshooting Steps:</strong>
                    <ol style="margin:10px 0;padding-left:20px;">
                        <li>Open browser console (F12) and check for detailed errors</li>
                        <li>Verify the field API name: <code>${ZOHO_SERVICE_API}</code></li>
                        <li>Check that you have permission to access Leads module</li>
                        <li>Ensure the widget has proper OAuth scopes</li>
                    </ol>
                </div>
            </div>`;
    }
});

ZOHO.embeddedApp.init();