/***********************
 * CONSTANTS & HELPERS
 ***********************/
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const formatMonth = (date) => `${MONTH_NAMES[date.getMonth()]} ${date.getFullYear()}`;
const getAllMonthsOfYear = (year) => MONTH_NAMES.map(m => `${m} ${year}`);


/*************************************
 * COQL DATA FETCH (OPTIMIZED)
 *************************************/
async function fetchFilteredLeadsCOQL(year) {

    const leadSources = [
        "Zoho Leads", "Zoho Partner", "Zoho CRM", "Zoho Partners 2024",
        "Zoho - Sutha", "Zoho - Hemanth", "Zoho - Sen", "Zoho - Audrey",
        "Zoho - Jacklyn", "Zoho - Adrian", "Zoho Partner Website", "Zoho - Chaitanya"
    ];

    const zohoServices = ["CRM", "CRMPlus", "One", "Bigin"];

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
            SELECT Created_Time, Lead_Source, Zoho_Service
            FROM Leads
            WHERE Lead_Source IN (${sourceList})
              AND Zoho_Service IN (${serviceList})
              AND Created_Time BETWEEN '${startDate}' AND '${endDate}'
            LIMIT ${offset}, ${limit}
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
 * GROUP LEADS BY MONTH & WEEK (UNCHANGED)
 ********************************************/
function groupLeadsByMonthWeek(leads, year) {

    const grouped = Object.fromEntries(
        getAllMonthsOfYear(year).map(m => [m, {1:0,2:0,3:0,4:0}])
    );

    leads.forEach(lead => {
        const createdTimeValue = lead.Created_Time;
        if (!createdTimeValue) return;

        try {
            const dateString = createdTimeValue.split(/[+Z]/)[0].replace("T"," ");
            const createdDate = new Date(dateString);

            if (isNaN(createdDate.getTime()) || createdDate.getFullYear() !== year) return;

            const monthKey = formatMonth(createdDate);
            if (!grouped[monthKey]) return;

            const week = Math.min(4, Math.max(1, Math.ceil(createdDate.getDate() / 7)));
            grouped[monthKey][week]++;

        } catch (e) {
            console.error("Date parsing error:", e);
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
 * TABLE RENDERING
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

    thead.innerHTML = `
        <tr style="background:#4CAF50;color:white;">
            <th style="padding:12px;border:1px solid #ddd;">Week</th>
            ${months.map(m => `<th style="padding:12px;border:1px solid #ddd;">${m}</th>`).join("")}
        </tr>
    `;

    for (let week = 1; week <= 4; week++) {

        const row = months.map((m, idx) => {
            const count = monthlyWeeklyCounts[m]?.[week] || 0;
            let prev = null;

            if (week > 1) {
                prev = monthlyWeeklyCounts[m]?.[week - 1] || 0;
            } else if (idx > 0) {
                const prevMonth = months[idx - 1];
                prev = monthlyWeeklyCounts[prevMonth]?.[4] || 0;
            }

            return `<td style="padding:10px;border:1px solid #ddd;text-align:center;">
                        ${count}${getPercentChange(count, prev)}
                    </td>`;
        }).join("");

        tbody.innerHTML += `
            <tr style="background:${week % 2 ? "#fff" : "#f9f9f9"};">
                <td style="padding:10px;border:1px solid #ddd;font-weight:bold;text-align:center;">Week ${week}</td>
                ${row}
            </tr>
        `;
    }

    let grandTotal = 0;

    const totalRow = months.map((m,i) => {
        const total = Object.values(monthlyWeeklyCounts[m]).reduce((a,b)=>a+b,0);
        grandTotal += total;

        const prev = i > 0
            ? Object.values(monthlyWeeklyCounts[months[i-1]]).reduce((a,b)=>a+b,0)
            : null;

        return `<td style="padding:12px;border:1px solid #ddd;text-align:center;">
                    <strong>${total}${getPercentChange(total, prev)}</strong>
                </td>`;
    }).join("");

    tbody.innerHTML += `
        <tr style="background:#e0e0e0;font-weight:bold;">
            <td style="padding:12px;border:1px solid #ddd;text-align:center;">Monthly Total</td>
            ${totalRow}
        </tr>
    `;

    document.querySelector("#footerNote").innerHTML = `
        <div style="background:#f5f5f5;padding:20px;border-radius:8px;">
            <strong>Filtered Leads:</strong> ${grandTotal}<br>
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
        const leads = await fetchFilteredLeadsCOQL(targetYear);

        if (!leads.length) {
            document.body.innerHTML = `
                <div style="padding:40px;text-align:center;">
                    <h2>No Matching Leads Found</h2>
                </div>
            `;
            return;
        }

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
