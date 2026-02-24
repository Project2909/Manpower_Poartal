const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(express.urlencoded({ extended: true }));

let storedData = null;

/* ---------------- UPLOAD PAGE ---------------- */

const UPLOAD_HTML = `
<!DOCTYPE html>
<html>
<head>
    <title>Company Manpower Allocation System</title>
    <style>
        body { font-family: 'Segoe UI'; background:#eef2f7; margin:0; }
        header { background:#0a3d62; color:white; padding:15px 40px; font-size:22px; }
        footer { background:#0a3d62; color:white; padding:10px; text-align:center; margin-top:40px; }
        .container { padding:40px; }
        .upload-box {
            background:white; padding:30px; width:50%;
            margin:auto; text-align:center; border-radius:8px;
        }
        button {
            background:#1e90ff; color:white; border:none;
            padding:12px 25px; margin-top:20px;
            cursor:pointer; font-size:16px;
        }
    </style>
</head>
<body>
<header>AADHVI TECHNOLOGIES – Manpower Allocation Portal</header>
<div class="container">
<div class="upload-box">
<form method="POST" enctype="multipart/form-data">
<h3>Upload Manpower Excel</h3>
<input type="file" name="excel" required><br><br>
<button type="submit">Upload</button>
</form>
</div>
</div>
<footer>© 2026 AADHVI Technologies | Manpower Planning System</footer>
</body>
</html>
`;

/* ---------------- ALLOCATION PAGE ---------------- */

function allocationPage(data) {
return `
<!DOCTYPE html>
<html>
<head>
<title>Manpower Allocation</title>
<style>
body { font-family:'Segoe UI'; background:#eef2f7; margin:0; }
header { background:#0a3d62; color:white; padding:15px 40px; font-size:22px; }
footer { background:#0a3d62; color:white; padding:10px; text-align:center; margin-top:40px; }
.container { padding:40px; }
table { border-collapse:collapse; width:100%; background:white; }
th,td { padding:12px; border:1px solid #ccc; text-align:center; }
th { background:#1e90ff; color:white; }
input[type=number]{ width:70px; }
button {
    background:#1e90ff; color:white; border:none;
    padding:12px 25px; margin-top:20px;
    cursor:pointer; font-size:16px;
}
</style>
</head>
<body>

<header>AADHVI TECHNOLOGIES – Manpower Allocation Portal</header>

<div class="container">
<form method="POST" action="/generate">
<table>
<tr>
<th>Select</th><th>Area</th><th>Current Manpower</th>
<th>Allocation %</th><th>Updated Manpower</th>
</tr>

${data.map((r,i)=>`
<tr>
<td><input type="checkbox" name="check_${i}"></td>
<td>${r.area}</td>
<td>${r.total}</td>
<td>
<input type="number" name="percent_${i}" value="100" min="0" max="100"
oninput="calc(${i})">
</td>
<td>
<input type="text" name="updated_${i}" value="${r.total}" readonly>
</td>
</tr>
`).join("")}

</table>
<input type="hidden" name="rows" value="${data.length}">
<button type="submit">Generate Updated Excel</button>
</form>
</div>

<footer>© 2026 AADHVI Technologies | Manpower Planning System</footer>

<script>
const data = ${JSON.stringify(data)};
function calc(i){
    let percent = document.getElementsByName("percent_"+i)[0].value;
    let total = data[i].total;
    document.getElementsByName("updated_"+i)[0].value =
        Math.round((total * percent) / 100);
}
</script>

</body>
</html>
`;
}

/* ---------------- ROUTES ---------------- */

app.get("/", (req,res)=>{
    res.send(UPLOAD_HTML);
});

app.post("/", upload.single("excel"), (req,res)=>{
    const workbook = XLSX.read(req.file.buffer);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    if(!json[0].Area || !json[0].Manpower){
        return res.send("Excel must contain Area and Manpower columns");
    }

    storedData = json;
    res.redirect("/allocate");
});

app.get("/allocate",(req,res)=>{
    const data = storedData.map(r=>({
        area: r.Area,
        total: r.Manpower
    }));
    res.send(allocationPage(data));
});

app.post("/generate",(req,res)=>{
    const rows = parseInt(req.body.rows);
    let output = [];

    for(let i=0;i<rows;i++){
        if(req.body["check_"+i]){
            let percent = parseInt(req.body["percent_"+i]);
            let original = storedData[i].Manpower;
            let updated = Math.round((original * percent)/100);

            output.push({
                Area: storedData[i].Area,
                Allocated_Percentage: percent,
                Updated_Manpower: updated
            });
        }
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(output);
    XLSX.utils.book_append_sheet(wb, ws, "Updated");

    const buffer = XLSX.write(wb,{ type:"buffer", bookType:"xlsx" });

    res.setHeader("Content-Disposition",
        "attachment; filename=Updated_Manpower_Selected_Areas.xlsx");
    res.send(buffer);
});

/* ---------------- RUN ---------------- */

app.listen(3000, ()=>{
    console.log("Server running on port 3000");
});
