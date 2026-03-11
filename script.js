let data = [];
let filtered = [];
let lastSelectedType = null;


// :small_blue_diamond: Fetch Data
fetch("companies.json")
    .then(res => res.json())
    .then(json => {
        data = json;
        filtered = data;
        loadFilters();
        updateKPIs();
        buildCityMidcSection();
        updateHorizontalChart(filtered);
        updateLineChart(filtered);
    });


// :small_blue_diamond: Helper → Check valid website
function isValidWebsite(url) {
    if (!url) return false;

    const val = url.toLowerCase().trim();
    return val !== "not available" && val !== "n/a" && val !== "na";
}


// :small_blue_diamond: Load Filters
function loadFilters() {
    let companySelect = document.getElementById("companyFilter");
    companySelect.innerHTML = '<option value="">Select Company</option>';

    data.forEach(d => {
        let opt = document.createElement("option");
        opt.value = d.company_name;
        opt.textContent = d.company_name;
        companySelect.appendChild(opt);
    });

    let types = [...new Set(data.map(d => d.company_based_on).filter(Boolean))];

    let typeSelect = document.getElementById("typeFilter");
    typeSelect.innerHTML = '<option value="">Select Type</option>';

    types.forEach(t => {
        let opt = document.createElement("option");
        opt.value = t;
        opt.textContent = t;
        typeSelect.appendChild(opt);
    });
}


// :small_blue_diamond: Event Listeners
document.getElementById("companyFilter").addEventListener("change", function () {

    applyFilters();

    let company = this.value.trim();

    if (company) {

        let selectedCompany = data.find(d =>
            (d.company_name || "").trim() === company
        );

        if (selectedCompany) {

            showDirectDetails(selectedCompany);

            // ⭐ Redirect to Details section
            document.getElementById("details").scrollIntoView({
                behavior: "smooth"
            });
        }
    }
});
document.getElementById("typeFilter").addEventListener("change", applyFilters);
document.getElementById("domainFilter").addEventListener("change", applyFilters);


// :small_blue_diamond: Apply Filters
function applyFilters() {
    let company = document.getElementById("companyFilter").value;
    let type = document.getElementById("typeFilter").value;
    let domain = document.getElementById("domainFilter").value;

    filtered = data.filter(d => {
        return (
            (!company || d.company_name === company) &&
            (!type || d.company_based_on === type) &&
            (!domain || (isValidWebsite(d.official_website) && d.official_website.includes(domain)))
        );
    });


    updateKPIs();
    buildCityMidcSection();
    updateHorizontalChart(filtered);
    updateLineChart(filtered); 

     // ⭐ NEW LOGIC
    if (company && filtered.length === 1) {
        showDirectDetails(filtered[0]);
    } else {
        document.getElementById("details").innerHTML =
            "<p>Please select a KPI card to view companies</p>";
    }
}
function getEmailDomain(email) {
    return email?.toLowerCase().trim().split("@")[1] || "";
}
   // :small_blue_diamond: KPI Calculation
function updateKPIs() {
    console.log("All Emails:", filtered.map(d => d.email));

    document.getElementById("total").textContent = filtered.length;

    document.getElementById("pvtltd").textContent =
    filtered.filter(d => {
        const name = (d.company_name || "")
            .toLowerCase()
            .replace(/\./g, "")   // remove dots
            .replace(/\s+/g, " ") // normalize spaces
            .trim();

        return /pvt\s*ltd$/.test(name) || name.endsWith("private limited");
    }).length;

    document.getElementById("llp").textContent =
    filtered.filter(d => {
        const name = (d.company_name || "")
            .toLowerCase()
            .replace(/\./g, "")   // remove dots
            .replace(/\s+/g, " ")
            .trim();

        return name.endsWith(" llp");
    }).length;

    document.getElementById("Pvt").textContent =
        filtered.filter(d => d.company_name?.toLowerCase().includes("Pvt")).length;

    document.getElementById("ltd").textContent =
    filtered.filter(d => {
        const name = (d.company_name || "")
        .toLowerCase()
        .replace(/\./g, "")   // remove dots
        .replace(/\s+/g, " ")
        .trim();
        return (
            (name.endsWith(" ltd") || name.endsWith(" limited")) &&
            !name.includes("pvt ltd") &&
            !name.includes("private limited") &&
            !name.includes("llp")
        );
    }).length;

    // :white_check_mark: FIXED WEBSITE LOGIC
    document.getElementById("no").textContent =
        filtered.filter(d => !isValidWebsite(d.official_website)).length;

    // 🌐 TOTAL WEBSITES KPI
    document.getElementById("totalWebsites").textContent =
        filtered.filter(d => isValidWebsite(d.official_website)).length;

    // :white_check_mark: DOMAIN KPIs (only valid websites)
    document.getElementById("com").textContent =
        filtered.filter(d => isValidWebsite(d.official_website) && d.official_website.includes(".com")).length;

    document.getElementById("in").textContent =
    filtered.filter(d => {
        const url = (d.official_website || "").toLowerCase();
        return (
            isValidWebsite(url) &&
            url.includes(".in") &&
            !url.endsWith(".co.in")
        );
    }).length;

    document.getElementById("org").textContent =
        filtered.filter(d => isValidWebsite(d.official_website) && d.official_website.includes(".org")).length;

    document.getElementById("io").textContent =
        filtered.filter(d => isValidWebsite(d.official_website) && d.official_website.includes(".io")).length;

    document.getElementById("coin").textContent =
    filtered.filter(d => {
        const url = (d.official_website || "").toLowerCase();
        return isValidWebsite(url) && url.endsWith(".co.in");
    }).length;

    document.getElementById("other").textContent =
    filtered.filter(d =>
        isValidWebsite(d.official_website) &&
        !d.official_website.includes(".com") &&
        !d.official_website.includes(".in") &&
        !d.official_website.includes(".org") &&
        !d.official_website.includes(".io") &&
        !d.official_website.includes(".co.in")
    ).length;

    // ⭐ COMPANY TYPE OTHER (exclude pvt ltd, llp, ltd)
    document.getElementById("otherType").textContent =
    filtered.filter(d => {
        const name = (d.company_name || "")
        .toLowerCase()
        .replace(/\./g, "")
        .replace(/\s+/g, " ")
        .trim();

    const isPvtLtd =
        /pvt\s*ltd$/.test(name) || name.endsWith("private limited");

    const isLLP = name.endsWith(" llp");

    const isLtd =
        (name.endsWith(" ltd") || name.endsWith(" limited")) &&
        !isPvtLtd &&
        !isLLP;

    return !isPvtLtd && !isLLP && !isLtd;
}).length;

    // 📧 TOTAL EMAILS
    document.getElementById("totalemail").textContent =
    filtered.filter(d => getEmailDomain(d.email || "")).length;

    // Gmail
    document.getElementById("gmail").textContent =
        filtered.filter(d => {
            let domain = getEmailDomain(d.email || "");
            return domain === "gmail.com";
        }).length;

    // Yahoo
    document.getElementById("yahoo").textContent =
        filtered.filter(d => {
            let domain = getEmailDomain(d.email || "");
            return domain === "yahoo.com";
        }).length;

    // Outlook / Hotmail
    document.getElementById("outlook").textContent =
        filtered.filter(d => {
            let domain = getEmailDomain(d.email || "");
            return domain === "outlook.com" || domain === "hotmail.com";
        }).length;

    // Company Email (not public providers)
    document.getElementById("companymail").textContent =
        filtered.filter(d => {
            let domain = getEmailDomain(d.email || "");
            return domain &&
                domain !== "gmail.com" &&
                domain !== "yahoo.com" &&
                domain !== "outlook.com" &&
                domain !== "hotmail.com";
        }).length;
}

function buildCityMidcSection(){

    const cityContainer = document.getElementById("cityCards");
    const midcContainer = document.getElementById("midcList");

    cityContainer.innerHTML = "";
    midcContainer.innerHTML = "";

    const cityCounts = {};
    const midcCounts = {};

    filtered.forEach(d => {

        let city = (d.city || "").trim();
        let midc = (d.MIDC || d.midc || "").trim().toLowerCase();

        if(city){
            cityCounts[city] = (cityCounts[city] || 0) + 1;
        }

        if(midc && midc.toLowerCase() !== "other"){
            midcCounts[midc] = (midcCounts[midc] || 0) + 1;
        }

    });

    /* CITY CARDS */

    const cities = Object.entries(cityCounts);
    
    
    let others = null;

    // Separate "Others"
    const normalCities = cities.filter(([city,count])=>{
        if(city.toLowerCase() === "others"){
            others = [city,count];
            return false;
        }
        return true;
    });

// Sort remaining cities
    normalCities.sort((a,b)=>b[1]-a[1]);

// Add "Others" at the end
    if(others){
        normalCities.push(others);
    }

    normalCities.forEach(([city,count])=>{

    let card = document.createElement("div");
    card.className = "city-card";

    if(city.toLowerCase() === "others"){
        card.classList.add("city-other");
    }

    card.innerHTML = `
        <div>
            <div class="city-name">${city}</div>
            <div class="city-count">${count} Companies</div>
        </div>
        <div>🏢</div>
    `;

    card.onclick = () => showCompaniesByCity(city);

    cityContainer.appendChild(card);

});


    /* MIDC LIST */

    Object.entries(midcCounts)
        .sort((a,b)=>b[1]-a[1])
        .slice(0,5)
        .forEach(([midc,count])=>{

        let item = document.createElement("div");
        item.className = "midc-item clickable-midc";
        
        item.innerHTML = `
        <div class="midc-name">${midc}</div>
        <div class="midc-count">${count}</div>
        
        `;
        
        
        item.onclick = () => showCompaniesByMidc(midc);
        midcContainer.appendChild(item);

    });

}
function renderCompanyList(list){

    let details = document.getElementById("details");

    if(list.length === 0){
        details.innerHTML = "<p>No companies found</p>";
        return;
    }

    let html = "";

    list.forEach(d => {

        html += `
        <div class="company-row">
            <span class="company-name">${d.company_name}</span>
            <button class="view-btn" onclick="viewFullDetails(${data.indexOf(d)})">
                View Details
            </button>
        </div>
        `;

    });

    details.innerHTML = html;

}

function showCompaniesByCity(city) {

    const list = filtered.filter(d => d.city === city);

    renderCompanyList(list);

    // ⭐ Scroll to details section
    document.getElementById("details").scrollIntoView({
        behavior: "smooth"
    });
}
 function showCompaniesByMidc(midc){

    const list = filtered.filter(d => {
        const zone = (d.MIDC || d.midc || d.Midc || "").trim().toLowerCase();
        return zone === midc.trim().toLowerCase();
    });

    renderCompanyList(list);

    // ⭐ Scroll to details section
    document.getElementById("details").scrollIntoView({
        behavior: "smooth"
    });
}
window.showCompanyList = function(type)  {
    console.log("Clicked:", type);
    lastSelectedType = type; 

    let details = document.getElementById("details");
    let list = [];

    switch(type) {

        case "totalWebsites":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website)
            );
            break;

        case "com":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".com")
            );
            break;

        case "in":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".in") &&
                !d.official_website.endsWith(".co.in")
            );
            break;

        case "org":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".org")
            );
            break;

        case "coin":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.endsWith(".co.in")
            );
            break;

        case "io":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".io")
            );
            break;

        case "yes":
            list = filtered.filter(d => isValidWebsite(d.official_website));
            break;

        case "no":
            list = filtered.filter(d => !isValidWebsite(d.official_website));
            break;

        case "pvtltd":
            list = filtered.filter(d => {
                const name = (d.company_name || "")
                .toLowerCase()
                .replace(/\./g, "")
                .replace(/\s+/g, " ")
                .trim();
                
                return /pvt\s*ltd$/.test(name) || name.endsWith("private limited");
            });
            break;

        case "llp":
            list = filtered.filter(d =>
                d.company_name?.toLowerCase().includes("llp")
            );
            break;

        case "ltd":
            list = filtered.filter(d => {
                let name = (d.company_name || "").toLowerCase().trim();
                return (
                    (name.endsWith("ltd") || name.endsWith("limited")) &&
                    !name.includes("pvt")
                );
            });
            break;
        
        case "otherType":
            list = filtered.filter(d => {
                const name = (d.company_name || "")
                .toLowerCase()
                .replace(/\./g, "")
                .replace(/\s+/g, " ")
                .trim();
                return (
                    !name.includes("pvt ltd") &&
                    !name.includes("private limited") &&
                    !name.endsWith(" llp") &&
                    !name.endsWith(" ltd") &&
                    !name.endsWith(" limited")
                );
            });
            break;

        case "other":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                !d.official_website.includes(".com") &&
                !d.official_website.includes(".in") &&
                !d.official_website.includes(".org") &&
                !d.official_website.includes(".io") &&
                !d.official_website.includes(".co.in")
            );
            break;

        case "totalemail":
            list = filtered.filter(d =>
                getEmailDomain(d.email || "")
            );
            break;

        case "gmail":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain === "gmail.com";
            });
            break;

        case "yahoo":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain === "yahoo.com";
            });
            break;

        case "outlook":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain === "outlook.com" || domain === "hotmail.com";
            });
            break;
  
        case "companymail":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain &&
                domain !== "gmail.com" &&
                domain !== "yahoo.com" &&
                domain !== "outlook.com" &&
                domain !== "hotmail.com";
            });
            break;
    }

    if (list.length === 0) {
        details.innerHTML = "<p>No companies found</p>";
        return;
    }

    let html = "";

    list.forEach((d, index) => {
        html += `
    <div class="company-row">
        <span class="company-name">${d.company_name}</span>
        <button class="view-btn" onclick="viewFullDetails(${data.indexOf(d)})">
            View Details
        </button>
    </div>
`;
    });

    details.innerHTML = html;
}
function formatWebsite(url) {
    if (!url.startsWith("http://") && !url.startsWith("https://")) {
        return "https://" + url;
    }
    return url;
}

function showDirectDetails(d) {

    let details = document.getElementById("details");

    details.innerHTML = `
<div class="full-details">
  <button class="back-btn mb-3" onclick="showCompanyList(lastSelectedType)">
        Back
    </button>
    <div class="detail-row">
        <div class="detail-label">Company Name</div>
        <div class="detail-value">${d.company_name}</div>
    </div>

    <div class="detail-row">
        <div class="detail-label">Website</div>
        <div class="detail-value">
            ${
                isValidWebsite(d.official_website)
                ? `<a href="${formatWebsite(d.official_website)}" target="_blank">${d.official_website}</a>`
                : "N/A"
            }
        </div>
    </div>

    <div class="detail-row">
        <div class="detail-label">Industry</div>
        <div class="detail-value">${d.company_based_on || ""}</div>
    </div>

    <div class="detail-row">
        <div class="detail-label">Goal</div>
        <div class="detail-value">${d.company_goal || ""}</div>
    </div>

    <div class="detail-row">
        <div class="detail-label">Address</div>
        <div class="detail-value">
            ${
                d.address
                ? `<a href="https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(d.address)}" target="_blank">${d.address}</a>`
                : ""
            }
        </div>
    </div>

    <div class="detail-row">
        <div class="detail-label">Mobile</div>
        <div class="detail-value">
            ${d.mobile_no ? `<a href="tel:${d.mobile_no}">${d.mobile_no}</a>` : ""}
        </div>
    </div>

    <div class="detail-row">
        <div class="detail-label">Email</div>
        <div class="detail-value">
            ${d.email ? `<a href="mailto:${d.email}">${d.email}</a>` : ""}
        </div>
    </div>

   

</div>
`;
}
window.viewFullDetails = function(index) {
    let d = data[index];
    showDirectDetails(d);
};

window.showCompanyList = function(type) {

    console.log("Clicked:", type);
    lastSelectedType = type;

    // Scroll to details section
    document.getElementById("details").scrollIntoView({
        behavior: "smooth"
    });

    let details = document.getElementById("details");
    let list = [];

    switch(type) {

        // ⭐ TOTAL COMPANIES
        case "total":
            list = filtered;
            break;

        case "totalWebsites":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website)
            );
            break;

        case "com":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".com")
            );
            break;

        case "in":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".in") &&
                !d.official_website.endsWith(".co.in")
            );
            break;

        case "org":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".org")
            );
            break;

        case "coin":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.endsWith(".co.in")
            );
            break;

        case "io":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                d.official_website.includes(".io")
            );
            break;

        case "yes":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website)
            );
            break;

        case "no":
            list = filtered.filter(d =>
                !isValidWebsite(d.official_website)
            );
            break;

        case "pvtltd":
            list = filtered.filter(d => {
                const name = (d.company_name || "")
                    .toLowerCase()
                    .replace(/\./g, "")
                    .replace(/\s+/g, " ")
                    .trim();

                return /pvt\s*ltd$/.test(name) || name.endsWith("private limited");
            });
            break;

        case "llp":
            list = filtered.filter(d =>
                d.company_name?.toLowerCase().includes("llp")
            );
            break;

        case "ltd":
            list = filtered.filter(d => {
                let name = (d.company_name || "").toLowerCase().trim();
                return (
                    (name.endsWith("ltd") || name.endsWith("limited")) &&
                    !name.includes("pvt")
                );
            });
            break;

        case "otherType":
            list = filtered.filter(d => {
                const name = (d.company_name || "")
                    .toLowerCase()
                    .replace(/\./g, "")
                    .replace(/\s+/g, " ")
                    .trim();

                return (
                    !name.includes("pvt ltd") &&
                    !name.includes("private limited") &&
                    !name.endsWith(" llp") &&
                    !name.endsWith(" ltd") &&
                    !name.endsWith(" limited")
                );
            });
            break;

        case "other":
            list = filtered.filter(d =>
                isValidWebsite(d.official_website) &&
                !d.official_website.includes(".com") &&
                !d.official_website.includes(".in") &&
                !d.official_website.includes(".org") &&
                !d.official_website.includes(".io") &&
                !d.official_website.includes(".co.in")
            );
            break;

        case "totalemail":
            list = filtered.filter(d =>
                getEmailDomain(d.email || "")
            );
            break;

        case "gmail":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain === "gmail.com";
            });
            break;

        case "yahoo":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain === "yahoo.com";
            });
            break;

        case "outlook":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain === "outlook.com" || domain === "hotmail.com";
            });
            break;

        case "companymail":
            list = filtered.filter(d => {
                let domain = getEmailDomain(d.email || "");
                return domain &&
                    domain !== "gmail.com" &&
                    domain !== "yahoo.com" &&
                    domain !== "outlook.com" &&
                    domain !== "hotmail.com";
            });
            break;
    }

    if (list.length === 0) {
        details.innerHTML = "<p>No companies found</p>";
        return;
    }

    let html = "";

    list.forEach(d => {
        html += `
        <div class="company-row">
            <span class="company-name">${d.company_name}</span>
            <button class="view-btn" onclick="viewFullDetails(${data.indexOf(d)})">
                View Details
            </button>
        </div>
        `;
    });

    details.innerHTML = html;
};
// sidear open/close
document.querySelectorAll(".icon-sidebar").forEach(icon => {
  icon.addEventListener("click", () => {
    document.getElementById("toggler").checked = false;
  });
});
// select option
const companyFilter = document.getElementById("companyFilter");

companies.forEach(company => {
  const option = document.createElement("option");
  option.value = company;
  option.textContent = company;
  companyFilter.appendChild(option);
});

function normalizeText(str) {
    return (str || "")
        .toLowerCase()
        .replace(/\s+/g, "")   // remove spaces
        .replace(/[^a-z0-9]/g, ""); // remove symbols
}

function runSearch(){

    const searchInput = document.getElementById("site-search");

    let rawValue = searchInput.value.toLowerCase().trim();
    let searchValue = normalizeText(rawValue);

    if(!rawValue) return;

    let list = [];

    // 1️⃣ DOMAIN SEARCH (.com .in etc)
    if(rawValue.startsWith(".")){

        list = data.filter(c =>
            isValidWebsite(c.official_website) &&
            c.official_website.toLowerCase().includes(rawValue)
        );

        renderCompanyList(list);

    }

    // 2️⃣ COMPANY NAME SEARCH
    else{

        let companyMatch = data.find(c =>
            normalizeText(c.company_name).includes(searchValue)
        );

        if(companyMatch){

            showDirectDetails(companyMatch);

        }else{

            // 3️⃣ INDUSTRY SEARCH
            list = data.filter(c =>
                normalizeText(c.company_based_on).includes(searchValue)
            );

            if(list.length > 0){

                renderCompanyList(list);

            }else{

                document.getElementById("details").innerHTML =
                    "<p>No results found</p>";

            }

        }

    }

    // ⭐ SCROLL DOWN
    setTimeout(() => {
        document.getElementById("details").scrollIntoView({
            behavior:"smooth"
        });
    },100);
}
document.getElementById("site-search").addEventListener("keydown", function(e){

    if(e.key === "Enter"){
        e.preventDefault();
        runSearch();
    }

});