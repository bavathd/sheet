let excelRows = [];
let headers = [];

window.onload = function () {
    loadExcel();
};

function loadExcel() {
    fetch("Paramount.xlsx")
        .then(res => res.arrayBuffer())
        .then(buffer => {
            const wb = XLSX.read(buffer, { type: "array" });
            const sheet = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            headers = json[0];
            excelRows = json.slice(1);

            renderList(excelRows);
        });
}

function renderList(rows) {
    const ul = document.getElementById("soldierList");
    ul.innerHTML = "";

    rows.forEach(row => {
        const name = row[1];
        const armyNo = row[4]; // REGT or Army No column (change if needed)

        const fileName = `${name.replace(/ /g, "")}_${armyNo}.pdf`;

        const li = document.createElement("li");
        li.textContent = `${name} (${armyNo})`;

        li.onclick = () => openDetails(row, fileName);

        ul.appendChild(li);
    });
}

function openDetails(row, file) {
    const rowObj = headers.map((h, i) => ({
        key: h,
        value: row[i] ?? ""
    }));

    window.location.href =
        `card.html?file=${file}&row=${encodeURIComponent(JSON.stringify(rowObj))}`;
}

function filterList() {
    const q = document.getElementById("search").value.toLowerCase();
    const filtered = excelRows.filter(r =>
        (r[1] + "").toLowerCase().includes(q) ||
        (r[4] + "").toLowerCase().includes(q)
    );
    renderList(filtered);
}
