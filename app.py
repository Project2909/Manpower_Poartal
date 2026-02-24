from flask import Flask, request, render_template_string, send_file, redirect, url_for, session
import pandas as pd
import io
import os   # ✅ REQUIRED FOR RENDER

app = Flask(__name__)
app.secret_key = "manpower-render-key"   # REQUIRED on Render

# ---------------- UPLOAD PAGE ----------------
UPLOAD_HTML = """ 
<!DOCTYPE html>
<html>
<head>
    <title>Company Manpower Allocation System</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #eef2f7; margin: 0; }
        header { background: #0a3d62; color: white; padding: 15px 40px; font-size: 22px; }
        footer { background: #0a3d62; color: white; padding: 10px; text-align: center; margin-top: 40px; }
        .container { padding: 40px; }
        .upload-box { background: white; padding: 30px; width: 50%; margin: auto; text-align: center; border-radius: 8px; }
        button { background: #1e90ff; color: white; border: none; padding: 12px 25px; margin-top: 20px; cursor: pointer; font-size: 16px; }
        button:hover { background: #0c6cd4; }
    </style>
</head>
<body>

<header>AADHVI TECHNOLOGIES – Manpower Allocation Portal</header>

<div class="container">
<div class="upload-box">
<form method="POST" enctype="multipart/form-data">
    <h3>Upload Manpower Excel</h3>
    <input type="file" name="excel" required>
    <br><br>
    <button type="submit">Upload</button>
</form>
</div>
</div>

<footer>© 2026 AADHVI Technologies | Manpower Planning System</footer>
</body>
</html>
"""

# ---------------- ALLOCATION PAGE ----------------
ALLOCATE_HTML = """ 
<!DOCTYPE html>
<html>
<head>
    <title>Manpower Allocation</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; background: #eef2f7; margin: 0; }
        header { background: #0a3d62; color: white; padding: 15px 40px; font-size: 22px; }
        footer { background: #0a3d62; color: white; padding: 10px; text-align: center; margin-top: 40px; }
        .container { padding: 40px; }
        table { border-collapse: collapse; width: 100%; background: white; }
        th, td { padding: 12px; border: 1px solid #ccc; text-align: center; }
        th { background: #1e90ff; color: white; }
        input[type=number] { width: 70px; }
        button { background: #1e90ff; color: white; border: none; padding: 12px 25px; margin-top: 20px; cursor: pointer; font-size: 16px; }
        button:hover { background: #0c6cd4; }
    </style>
</head>

<body>
<header>AADHVI TECHNOLOGIES – Manpower Allocation Portal</header>

<div class="container">
<form method="POST" action="/generate">
<table>
<tr>
    <th>Select</th>
    <th>Area</th>
    <th>Current Manpower</th>
    <th>Allocation %</th>
    <th>Updated Manpower</th>
</tr>

{% for row in data %}
<tr>
    <td><input type="checkbox" name="check_{{ loop.index0 }}"></td>
    <td>{{ row.area }}</td>
    <td>{{ row.total }}</td>
    <td>
        <input type="number" name="percent_{{ loop.index0 }}" value="100" min="0" max="100"
               oninput="calc({{ loop.index0 }})">
    </td>
    <td>
        <input type="text" name="updated_{{ loop.index0 }}" value="{{ row.total }}" readonly>
    </td>
</tr>
{% endfor %}
</table>

<input type="hidden" name="rows" value="{{ data|length }}">
<button type="submit">Generate Updated Excel</button>
</form>

<script>
const data = {{ data|tojson }};
function calc(i){
    let percent = document.getElementsByName("percent_" + i)[0].value;
    let total = data[i].total;
    let updated = Math.round((total * percent) / 100);
    document.getElementsByName("updated_" + i)[0].value = updated;
}
</script>
</div>

<footer>© 2026 AADHVI Technologies | Manpower Planning System</footer>
</body>
</html>
"""

# ---------------- ROUTES ----------------

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files["excel"]
        df = pd.read_excel(file)   # ✅ works after openpyxl install

        if "Area" not in df.columns or "Manpower" not in df.columns:
            return "Excel must contain Area and Manpower columns"

        session["stored_df"] = df.to_json()
        return redirect(url_for("allocate"))

    return render_template_string(UPLOAD_HTML)


@app.route("/allocate")
def allocate():
    if "stored_df" not in session:
        return redirect(url_for("upload"))

    stored_df = pd.read_json(session["stored_df"])
    data = [{"area": r["Area"], "total": r["Manpower"]} for _, r in stored_df.iterrows()]
    return render_template_string(ALLOCATE_HTML, data=data)


@app.route("/generate", methods=["POST"])
def generate():
    stored_df = pd.read_json(session["stored_df"])
    rows = int(request.form["rows"])
    output = []

    for i in range(rows):
        if f"check_{i}" in request.form:
            percent = int(request.form[f"percent_{i}"])
            original = stored_df.at[i, "Manpower"]
            updated = round((original * percent) / 100)

            output.append({
                "Area": stored_df.at[i, "Area"],
                "Allocated_Percentage": percent,
                "Updated_Manpower": updated
            })

    out_df = pd.DataFrame(output)
    buffer = io.BytesIO()
    out_df.to_excel(buffer, index=False)
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="Updated_Manpower_Selected_Areas.xlsx"
    )

# ---------------- RUN (RENDER SAFE) ----------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # ✅ FIX
    app.run(host="0.0.0.0", port=port)        # ✅ FIX
