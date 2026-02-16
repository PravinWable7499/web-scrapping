let companies = [];

const dropdown = document.getElementById("companySelect");
const container = document.getElementById("companyDetails");
const loader = document.getElementById("loader");

/* ===============================
   FETCH EXCEL DATA
================================= */
fetch("companies.xlsx")
.then(res => res.arrayBuffer())
.then(data => {

    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    companies = XLSX.utils.sheet_to_json(worksheet);
    populateDropdown();
})
.catch(error => {
    console.error("Error loading Excel file:", error);
});


/* ===============================
   POPULATE DROPDOWN
================================= */
function populateDropdown() {

    dropdown.innerHTML = "";

    let defaultOption = document.createElement("option");
    defaultOption.text = "Search and Select Company";
    defaultOption.value = "";
    dropdown.appendChild(defaultOption);

    companies.forEach((company, index) => {
        let option = document.createElement("option");
        option.value = index;
        option.text = company.company_name;
        dropdown.appendChild(option);
    });

    $('#companySelect').select2({
        placeholder: "Search and Select Company",
        allowClear: true
    });

    $('#companySelect').on('change', function () {
        showCompanyData(this.value);
    });
}


/* ===============================
   SHOW COMPANY DETAILS
================================= */
function showCompanyData(index) {

    container.innerHTML = "";
    container.style.display = "none";

    if (index === "") return;

    loader.style.display = "block";

    setTimeout(() => {

        loader.style.display = "none";

        const company = companies[index];

        for (let key in company) {

            let row = document.createElement("div");
            row.className = "row";

            let keyDiv = document.createElement("div");
            keyDiv.className = "key";
            keyDiv.innerText = formatKey(key);

            let valueDiv = document.createElement("div");
            valueDiv.className = "value";

            let value = company[key];

            /* ===== WEBSITE ===== */
            if (typeof value === "string" && value.startsWith("http")) {
                valueDiv.innerHTML = `<a href="${value}" target="_blank">${value}</a>`;
            }

            /* ===== EMAIL ===== */
            else if (key.toLowerCase().includes("email")) {
                valueDiv.innerHTML = `<a href="mailto:${value}">${value}</a>`;
            }

            /* ===== ADDRESS → GOOGLE MAPS ===== */
            else if (key.toLowerCase().includes("address")) {
                let mapLink = `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(value)}`;
                valueDiv.innerHTML = `<a href="${mapLink}" target="_blank">${value}</a>`;
            }

            /* ===== MOBILE → DIRECT CALL ===== */
            else if (key.toLowerCase().includes("mobile")) {
                valueDiv.innerHTML = `<a href="tel:${value}">${value}</a>`;
            }

            /* ===== NORMAL TEXT ===== */
            else {
                valueDiv.innerText = value;
            }

            row.appendChild(keyDiv);
            row.appendChild(valueDiv);
            container.appendChild(row);
        }

        container.style.display = "block";

    }, 400);
}


/* ===============================
   FORMAT COLUMN NAMES
================================= */
function formatKey(key) {

    const lowerKey = key.toLowerCase();

    // Custom Renaming
    if (lowerKey === "company_goal") return "About Company";
    if (lowerKey === "company_based_on") return "Based On";
    if (lowerKey === "mobile_no") return "Mobile No";
    if (lowerKey === "company_name") return "Company Name";

    // Default Formatting
    return key
        .replace(/_/g, " ")
        .replace(/\b\w/g, l => l.toUpperCase());
}
