/*******************************
 * CONSTANTS & HELPERS
 *******************************/
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const formatMonth = (date) => `${MONTH_NAMES[date.getMonth()]} ${date.getFullYear()}`;
const getAllMonthsOfYear = (year) => MONTH_NAMES.map(m => `${m} ${year}`);

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
 * FETCH FILTERED LEADS USING COQL
 *************************************/
async function fetchFilteredLeadsCOQL() {
    const year = new Date().getFullYear();

    const leadSources = [
        "Zoho Leads", "Zoho Partner", "Zoho CRM", "Zoho Partners 2024",
        "Zoho - Sutha", "Zoho - Hemanth", "Zoho - Sen", "Zoho - Audrey",
        "Zoho - Jacklyn", "Zoho - Adrian", "Zoho Partner Website", "Zoho - Chaitanya"
    ];
    const zohoServices = ["Desk", "Workplace", "Projects", "Mail"];

    const startDate = `${year}-01-01T00:00:00+00:00`;
    const endDate   = `${year}-12-31T23:59:59+00:00`;

    const sourceList  = leadSources.map(v => `'${v}'`).join(",");
    const serviceList = zohoServices.map(v => `'${v}'`).join(",");

    const query = `
        SELECT Created_Time, Lead_Source, Zoho_Service
        FROM Leads
        WHERE Lead_Source IN (${sourceList})
          AND Zoho_Service IN (${serviceList})
          AND Created_Time BETWEEN '${startDate}' AND '${endDate}'
    `;

    try {
        const resp = await ZOHO.CRM.API.coql({ select_query: query });
        return resp?.data || [];
    } catch (err) {
        console.error("Error fetching leads via COQL:", err);
        return [];
    }
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

        const date = new Date(createdTime);
        if (isNaN(date.getTime()) || date.getFullYear() !== year) return;

        const monthKey = formatMonth(date);
        const week = getWeekOfMonth(date);
        if (grouped[monthKey]?.[week] !== undefined) grouped[monthKey][week]++;
    });

    return grouped;
}

/*************************************
 * PERCENT CHANGE HELPER
 *************************************/
function getPercentChange(current, previous) {
    if (previous == null) return "";
    if (previous === 0) return current > 0
        ? ` <span style="color:green;font-weight:bold;">(+∞)</span>`
        : ` <span style="color:gray;">(0%)</span>`;
    const change = (((current - previous) / previous) * 100).toFixed(1);
    const color = change > 0 ? "green" : change < 0 ? "red" : "gray";
    return ` <span style="color:${color};font-weight:bold;">(${change > 0 ? "+" : ""}${change}%)</span>`;
}

/*************************************
 * TABLE RENDERING (DYNAMIC WEEKS)
 *************************************/
function renderTable(monthlyWeeklyCounts, year, totalFiltered) {
    let table = document.querySelector("#leadsTable");
    if (!table) {
        document.body.innerHTML += `
            <div style="margin:20px;">
                <h2>Lead Generation Report - ${year}</h2>
                <table id="leadsTable" style="border-collapse:collapse;width:100%;margin:20px 0;">
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
    const maxWeeks = Math.max(...months.map(m => Object.keys(monthlyWeeklyCounts[m]).length));

    // Table Header
    thead.innerHTML = `<tr style="background:#4CAF50;color:white;">
        <th style="padding:12px;border:1px solid #ddd;text-align:center;">Week</th>
        ${months.map(m => `<th style="padding:12px;border:1px solid #ddd;text-align:center;">${m}</th>`).join("")}
    </tr>`;

    // Week Rows
    for (let week = 1; week <= maxWeeks; week++) {
        const row = months.map((m, idx) => {
            const count = monthlyWeeklyCounts[m]?.[week] ?? "";
            let prev = null;
            if (week > 1) prev = monthlyWeeklyCounts[m]?.[week - 1] ?? null;
            else if (week === 1 && idx > 0) prev = monthlyWeeklyCounts[months[idx - 1]]?.[maxWeeks] ?? null;
            return `<td style="padding:10px;border:1px solid #ddd;text-align:center;">${count !== "" ? count + getPercentChange(count, prev) : ""}</td>`;
        }).join("");
        tbody.innerHTML += `<tr style="background:${week % 2 ? "white" : "#f9f9f9"};">
            <td style="padding:10px;border:1px solid #ddd;font-weight:bold;text-align:center;">Week ${week}</td>
            ${row}
        </tr>`;
    }

    // Monthly Totals
    let grandTotal = 0;
    const totalRow = months.map((m,i) => {
        const total = Object.values(monthlyWeeklyCounts[m] || {}).reduce((a,b)=>a+b,0);
        grandTotal += total;
        const prev = i>0 ? Object.values(monthlyWeeklyCounts[months[i-1]] || {}).reduce((a,b)=>a+b,0) : null;
        return `<td style="padding:12px;border:1px solid #ddd;text-align:center;"><strong>${total}${getPercentChange(total, prev)}</strong></td>`;
    }).join("");
    tbody.innerHTML += `<tr style="background:#f0f1f2;color:white;font-weight:bold;">
        <td style="padding:12px;border:1px solid #ddd;text-align:center;">Monthly Total</td>${totalRow}
    </tr>`;

    document.querySelector("#footerNote").innerHTML = `
        <div style="background:#f5f5f5;padding:20px;border-radius:8px;margin-top:20px;">
            <h4 style="margin:0 0 15px;color:#333;font-size:0.9em">Lead Generation Summary for ${year}</h4>
            <div style="display:flex;flex-wrap:wrap;gap:20px;font-size:0.9em">
                <div><strong>Total Leads:</strong> <span style="color:#2196F3;font-size:1.2em;">${grandTotal}</span></div>
                <div><strong>Total Fetched:</strong> ${totalFiltered}</div>
                <div><strong>Report Period:</strong> Jan - Dec ${year}</div>
                <div><strong>Generated:</strong> ${new Date().toLocaleString()}</div>
            </div>
            <p style="margin-top:15px;color:#666;font-size:0.9em;">
                Week-over-week and month-over-month percentage changes shown.
            </p>
        </div>`;
}

/*************************************
 * PAGE LOAD
 *************************************/
ZOHO.embeddedApp.on("PageLoad", async () => {
    const year = new Date().getFullYear();

    document.body.innerHTML = `
        <div id="loadingDiv" style="text-align:center;padding:40px;background:#e3f2fd;border-radius:8px;margin:20px;">
            <h2>Loading Lead Data for ${year}</h2>
            <p>Applying Lead Source and Zoho Service filters...</p>
        </div>
        <table id="leadsTable" style="border-collapse:collapse;width:100%;margin:20px 0;"><thead></thead><tbody></tbody></table>
        <div id="footerNote"></div>`;

    try {
        const leads = await fetchFilteredLeadsCOQL();

        if (!leads.length) {
            document.body.innerHTML = `<div style="text-align:center;padding:40px;background:#fff3cd;border-radius:8px;margin:20px;border-left:5px solid #ffc107;">
                <h2>No Matching Leads Found</h2>
                <p>No leads found matching the specified Lead Source and Zoho Service criteria.</p>
            </div>`;
            return;
        }

        const grouped = groupLeadsByMonthWeek(leads, year);
        renderTable(grouped, year, leads.length);

        document.getElementById("loadingDiv").style.display = "none";
    } catch(err) {
        document.body.innerHTML = `<div style="text-align:center;padding:40px;background:#f8d7da;border-radius:8px;margin:20px;border-left:5px solid #dc3545;">
            <h2>Error Loading Report</h2>
            <p><strong>${err.message}</strong></p>
            <button onclick="location.reload()" style="padding:10px 20px;background:#007bff;color:white;border:none;border-radius:5px;cursor:pointer;">
                Retry
            </button>
        </div>`;
    }
});

ZOHO.embeddedApp.init();