/***********************
 * CONSTANTS & HELPERS
 ***********************/
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const formatMonth = (date) => `${MONTH_NAMES[date.getMonth()]} ${date.getFullYear()}`;
const getAllMonthsOfYear = (year) => MONTH_NAMES.map(m => `${m} ${year}`);

/**
 * Returns number of weeks in a month (4 or 5)
 */
function getWeeksInMonth(year, monthIndex) {
    const firstDay = new Date(year, monthIndex, 1).getDay();
    const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
    return Math.ceil((firstDay + daysInMonth) / 7);
}

/**
 * Returns week number (1–5) within the month
 */
function getWeekOfMonth(date) {
    const firstDay = new Date(date.getFullYear(), date.getMonth(), 1).getDay();
    return Math.ceil((date.getDate() + firstDay) / 7);
}


/*************************************
 * COQL DATA FETCH (CURRENT YEAR ONLY)
 *************************************/
async function fetchFilteredLeadsCOQL() {

    const year = new Date().getFullYear();

    const leadSources = [
        "Zoho Leads","Zoho Partner","Zoho CRM","Zoho Partners 2024",
        "Zoho - Sutha","Zoho - Hemanth","Zoho - Sen","Zoho - Audrey",
        "Zoho - Jacklyn","Zoho - Adrian","Zoho Partner Website","Zoho - Chaitanya"
    ];

    const zohoServices = ["CRM","CRMPlus","One","Bigin"];

    const startDate = `${year}-01-01T00:00:00+00:00`;
    const endDate   = `${year}-12-31T23:59:59+00:00`;

    const sourceList  = leadSources.map(v => `'${v}'`).join(",");
    const serviceList = zohoServices.map(v => `'${v}'`).join(",");

    let allData = [];
    let offset = 0;
    const limit = 2000;
    let hasMore = true;

    while (hasMore) {

        const query = `
            SELECT Created_Time
            FROM Leads
            WHERE Lead_Source IN (${sourceList})
              AND Zoho_Service IN (${serviceList})
              AND Created_Time BETWEEN '${startDate}' AND '${endDate}'
            LIMIT ${limit} OFFSET ${offset}
        `;

        const resp = await ZOHO.CRM.API.coql({ select_query: query });
        const data = resp?.data || [];

        allData.push(...data);
        hasMore = data.length === limit;
        offset += limit;
    }

    return allData;
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
        if (!lead.Created_Time) return;

        const date = new Date(lead.Created_Time);
        if (isNaN(date) || date.getFullYear() !== year) return;

        const monthKey = formatMonth(date);
        const week = getWeekOfMonth(date);

        if (grouped[monthKey]?.[week] !== undefined) {
            grouped[monthKey][week]++;
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
            : ` <span style="color:gray;">(0%)</span>`;
    }

    const change = (((current - previous) / previous) * 100).toFixed(1);
    const color = change > 0 ? "green" : change < 0 ? "red" : "gray";

    return ` <span style="color:${color};font-weight:bold;">(${change > 0 ? "+" : ""}${change}%)</span>`;
}


/*************************************
 * TABLE RENDERING (DYNAMIC WEEKS)
 *************************************/
function renderTable(monthlyWeeklyCounts, year, totalFiltered) {

    document.body.innerHTML = `
        <div style="margin:20px;">
            <h2>Lead Generation Report - ${year}</h2>
            <table id="leadsTable" style="border-collapse:collapse;width:100%;margin:20px 0;">
                <thead></thead>
                <tbody></tbody>
            </table>
            <div id="footerNote"></div>
        </div>
    `;

    const table = document.querySelector("#leadsTable");
    const thead = table.querySelector("thead");
    const tbody = table.querySelector("tbody");

    const months = getAllMonthsOfYear(year);
    const maxWeeks = 5;

    thead.innerHTML = `
        <tr style="background:#4CAF50;color:white;">
            <th>Week</th>
            ${months.map(m => `<th>${m}</th>`).join("")}
        </tr>
    `;

    for (let week = 1; week <= maxWeeks; week++) {

        const row = months.map((m, idx) => {
            const count = monthlyWeeklyCounts[m]?.[week] ?? "";
            let prev = null;

            if (week > 1) prev = monthlyWeeklyCounts[m]?.[week - 1] ?? null;
            else if (idx > 0) prev = monthlyWeeklyCounts[months[idx - 1]]?.[5] ?? null;

            return `<td style="text-align:center;">
                        ${count !== "" ? count + getPercentChange(count, prev) : ""}
                    </td>`;
        }).join("");

        tbody.innerHTML += `
            <tr>
                <td style="font-weight:bold;text-align:center;">Week ${week}</td>
                ${row}
            </tr>
        `;
    }

    document.querySelector("#footerNote").innerHTML = `
        <div style="background:#f5f5f5;padding:20px;border-radius:8px;">
            <strong>Total Records Retrieved:</strong> ${totalFiltered}<br>
            <strong>Period:</strong> Jan–Dec ${year}<br>
            <strong>Generated:</strong> ${new Date().toLocaleString()}
        </div>
    `;
}


/*************************************
 * PAGE LOAD
 *************************************/
ZOHO.embeddedApp.on("PageLoad", async () => {

    const targetYear = new Date().getFullYear();

    document.body.innerHTML = `
        <div style="text-align:center;padding:40px;">
            <h2>Loading Lead Report for ${targetYear}...</h2>
            <p>Applying COQL filters...</p>
        </div>
    `;

    try {
        const leads = await fetchFilteredLeadsCOQL();
        const grouped = groupLeadsByMonthWeek(leads, targetYear);
        renderTable(grouped, targetYear, leads.length);

    } catch (err) {
        document.body.innerHTML = `
            <div style="padding:40px;text-align:center;color:red;">
                <h2>Error Loading Report</h2>
                <pre>${err.message}</pre>
            </div>
        `;
    }
});

ZOHO.embeddedApp.init();
