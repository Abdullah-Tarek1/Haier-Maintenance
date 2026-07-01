// ==========================
// Haier Maintenance System
// ==========================
const BREAKDOWN_FORM_URL = "https://forms.office.com/r/X6jwzcSjbs";
const PM_FORM_URL = "https://forms.office.com/r/REPLACE_WITH_PM_FORM";
const EXCEL_FILE_URL = "https://haiergroup-my.sharepoint.com/:x:/g/personal/hee_abdullah_elbadawy_haier_com/IQDaQWGew4WlTag6vFJSXK-2AZfz0IFtIfwjFSs2Epg30L0?e=3728PT";

const machines = {
    "HEE-M01": "ماكينة تشكيل الكابينة البلاستيك",
    "HEE-M02": "ماكينة تشكيل الباب البلاستيك",
    "HEE-M03": "ماكينة تشكيل الكابينة الصاج",
    "HEE-M04": "ماكينة تشكيل الباب الصاج",
    "HEE-M05": "ماكينة حقن الكابينة",
    "HEE-M06": "ماكينة حقن الباب",
    "HEE-M07": "طلمبات الفاكيوم",
    "HEE-M08": "ماكينة شحن الفريون",
    "HEE-M09": "ماكينة الKM",
    "HEE-M10": "خط التجميع والتغليف",
    "HEE-M11": "ماكينة حقن البلاستيك",
    "HEE-M12": "ماكينة سحب الشيت البلاستيك"
};

const BREAKDOWN_FORMS_BY_MACHINE = {
    "HEE-M01": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%AA%D8%B4%D9%83%D9%8A%D9%84%20%D8%A7%D9%84%D9%83%D8%A7%D8%A8%D9%8A%D9%86%D8%A9%20%D8%A7%D9%84%D8%A8%D9%84%D8%A7%D8%B3%D8%AA%D9%8A%D9%83%22",
    "HEE-M02": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%AA%D8%B4%D9%83%D9%8A%D9%84%20%D8%A7%D9%84%D8%A8%D8%A7%D8%A8%20%D8%A7%D9%84%D8%A8%D9%84%D8%A7%D8%B3%D8%AA%D9%8A%D9%83%22",
    "HEE-M03": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%AA%D8%B4%D9%83%D9%8A%D9%84%20%D8%A7%D9%84%D9%83%D8%A7%D8%A8%D9%8A%D9%86%D8%A9%20%D8%A7%D9%84%D8%B5%D8%A7%D8%AC%22",
    "HEE-M04": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%AA%D8%B4%D9%83%D9%8A%D9%84%20%D8%A7%D9%84%D8%A8%D8%A7%D8%A8%20%D8%A7%D9%84%D8%B5%D8%A7%D8%AC%22",
    "HEE-M05": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%AD%D9%82%D9%86%20%D8%A7%D9%84%D9%83%D8%A7%D8%A8%D9%8A%D9%86%D8%A9%22",
    "HEE-M06": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%AD%D9%82%D9%86%20%D8%A7%D9%84%D8%A8%D8%A7%D8%A8%22",
    "HEE-M07": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D8%B7%D9%84%D9%85%D8%A8%D8%A7%D8%AA%20%D8%A7%D9%84%D9%81%D8%A7%D9%83%D9%8A%D9%88%D9%85%22",
    "HEE-M08": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%B4%D8%AD%D9%86%20%D8%A7%D9%84%D9%81%D8%B1%D9%8A%D9%88%D9%86%22",
    "HEE-M09": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%A7%D9%84KM%22",
    "HEE-M10": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D8%AE%D8%B7%20%D8%A7%D9%84%D8%AA%D8%AC%D9%85%D9%8A%D8%B9%20%D9%88%D8%A7%D9%84%D8%AA%D8%BA%D9%84%D9%8A%D9%81%22",
    "HEE-M11": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%AD%D9%82%D9%86%20%D8%A7%D9%84%D8%A8%D9%84%D8%A7%D8%B3%D8%AA%D9%8A%D9%83%22",
    "HEE-M12": "https://forms.office.com/Pages/ResponsePage.aspx?id=tYX4NXiumUqB_4DmmpzBod6AiZhUeDtBrRQQcOKU8OZUNTZCTVZOQ1ExRFBJNDJWWjQ1WDAzTlVWTy4u&r67953c89bedd41f580798fea7849e29c=%22%D9%85%D8%A7%D9%83%D9%8A%D9%86%D8%A9%20%D8%B3%D8%AD%D8%A8%20%D8%A7%D9%84%D8%B4%D9%8A%D8%AA%20%D8%A7%D9%84%D8%A8%D9%84%D8%A7%D8%B3%D8%AA%D9%8A%D9%83%22"
};

function openBreakdownForm(machineId) {
    const url = BREAKDOWN_FORMS_BY_MACHINE[machineId] || BREAKDOWN_FORM_URL;
    window.open(url, "_blank");
}

function openPMForm(machineId) {
    if (PM_FORM_URL.includes("REPLACE")) {
        alert("⚠️ لم يتم إعداد نموذج الصيانة الوقائية بعد");
        return;
    }
    window.open(PM_FORM_URL, "_blank");
}

function openExcel() {
    window.open(EXCEL_FILE_URL, "_blank");
}

function getMachineIdFromUrl() {
    const params = new URLSearchParams(window.location.search);
    return params.get("id");
}

document.addEventListener("DOMContentLoaded", function() {
    // machine.html
    const machineId = getMachineIdFromUrl();
    if (machineId && machines[machineId]) {
        const idEl = document.getElementById("machineId");
        const nameEl = document.getElementById("machineName");
        if (idEl) idEl.textContent = machineId;
        if (nameEl) nameEl.textContent = machines[machineId];

        const btnBD = document.getElementById("btnBreakdown");
        const btnPM = document.getElementById("btnPM");
        const btnEx = document.getElementById("btnExcel");
        if (btnBD) btnBD.onclick = () => openBreakdownForm(machineId);
        if (btnPM) btnPM.onclick = () => openPMForm(machineId);
        if (btnEx) btnEx.onclick = openExcel;
    }

    // report.html - populate dropdown
    const sel = document.getElementById("machineSelect");
    if (sel) {
        for (const [id, name] of Object.entries(machines)) {
            const opt = document.createElement("option");
            opt.value = id;
            opt.textContent = id + " - " + name;
            sel.appendChild(opt);
        }
    }

    const btnReport = document.getElementById("btnReportBreakdown");
    if (btnReport) {
        btnReport.onclick = () => {
            const s = document.getElementById("machineSelect");
            openBreakdownForm(s ? s.value : "");
        };
    }

    // machines.html - build grid
    const grid = document.getElementById("machinesGrid");
    if (grid) {
        for (const [id, name] of Object.entries(machines)) {
            const card = document.createElement("a");
            card.href = "machine.html?id=" + id;
            card.className = "machine-card";
            card.innerHTML = '<div class="machine-id-badge">' + id + '</div><div class="machine-name">' + name + '</div>';
            grid.appendChild(card);
        }
    }

    // dashboard.html - Excel button
    const btnEx = document.getElementById("btnOpenExcel");
    if (btnEx) btnEx.onclick = openExcel;
});
