from flask import Flask, render_template_string, request, send_file
from openpyxl import load_workbook, Workbook
import csv
import io

app = Flask(__name__)

HTML_TEMPLATE = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Vertikalios ataskaitos kepėja</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: #f0f2f5;
      display: flex; justify-content: center; align-items: center;
      min-height: 100vh; padding: 20px;
    }
    .card {
      background: #fff; border-radius: 12px; padding: 40px;
      box-shadow: 0 2px 16px rgba(0,0,0,0.08);
      max-width: 520px; width: 100%; text-align: center;
    }
    h1 { font-size: 1.3rem; color: #1a1a2e; margin-bottom: 8px; }
    .subtitle { color: #666; font-size: 0.9rem; margin-bottom: 24px; }
    .drop-zone {
      border: 2px dashed #c0c6d0; border-radius: 10px;
      padding: 40px 20px; cursor: pointer;
      transition: all 0.2s ease; background: #fafbfc;
    }
    .drop-zone:hover, .drop-zone.drag-over {
      border-color: #4a6cf7; background: #eef1ff;
    }
    .drop-zone svg { width: 48px; height: 48px; color: #4a6cf7; margin-bottom: 12px; }
    .drop-zone p { color: #555; font-size: 0.95rem; }
    .drop-zone .hint { color: #999; font-size: 0.8rem; margin-top: 6px; }
    .file-name {
      margin-top: 16px; padding: 10px 16px; background: #eef1ff;
      border-radius: 8px; color: #4a6cf7; font-size: 0.9rem;
      display: none; align-items: center; justify-content: center; gap: 8px;
    }
    .file-name .remove {
      cursor: pointer; color: #e74c3c; font-weight: bold; font-size: 1.1rem;
    }
    input[type="file"] { display: none; }
    button[type="submit"] {
      margin-top: 20px; padding: 12px 32px; font-size: 1rem;
      background: #4a6cf7; color: #fff; border: none; border-radius: 8px;
      cursor: pointer; transition: background 0.2s; width: 100%;
    }
    button[type="submit"]:hover { background: #3a5ce5; }
    button[type="submit"]:disabled { background: #b0b8c9; cursor: not-allowed; }
  </style>
</head>
<body>
  <div class="card">
    <h1>Vertikalios ataskaitos kepėja</h1>
    <p class="subtitle">Įkelkite duomenis iš VDA sistemos</p>
    <form action="/" method="post" enctype="multipart/form-data" id="uploadForm">
      <div class="drop-zone" id="dropZone">
        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="1.5">
          <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5m-13.5-9L12 3m0 0l4.5 4.5M12 3v13.5"/>
        </svg>
        <p>Nutempkite failą čia arba paspauskite</p>
        <p class="hint">.xlsx arba .csv formatai</p>
      </div>
      <input type="file" name="source_file" id="fileInput" accept=".xlsx,.csv" required>
      <div class="file-name" id="fileName">
        <span id="fileNameText"></span>
        <span class="remove" id="removeFile">&times;</span>
      </div>
      <button type="submit" id="submitBtn" disabled>Atsisiųsti ataskaitą</button>
    </form>
  </div>
  <script>
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const fileNameText = document.getElementById('fileNameText');
    const removeFile = document.getElementById('removeFile');
    const submitBtn = document.getElementById('submitBtn');

    function showFile(file) {
      fileNameText.textContent = file.name;
      fileName.style.display = 'flex';
      submitBtn.disabled = false;
    }
    function clearFile() {
      fileInput.value = '';
      fileName.style.display = 'none';
      submitBtn.disabled = true;
    }

    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', () => {
      if (fileInput.files.length) showFile(fileInput.files[0]);
    });

    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('drag-over');
      const file = e.dataTransfer.files[0];
      if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.csv'))) {
        fileInput.files = e.dataTransfer.files;
        showFile(file);
      }
    });
    removeFile.addEventListener('click', (e) => { e.stopPropagation(); clearFile(); });
  </script>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "source_file" not in request.files:
            return "Failas neįkeltas", 400

        source_file = request.files["source_file"]
        filename = source_file.filename.lower()

        try:
            if filename.endswith('.csv'):
                source_wb = Workbook()
                vda_source = source_wb.active
                text = source_file.read().decode('utf-8')
                reader = csv.reader(io.StringIO(text))
                for row_idx, row_data in enumerate(reader, start=1):
                    for col_idx, value in enumerate(row_data, start=1):
                        if value.strip() == '':
                            vda_source.cell(row=row_idx, column=col_idx).value = None
                        else:
                            try:
                                vda_source.cell(row=row_idx, column=col_idx).value = float(value)
                            except ValueError:
                                vda_source.cell(row=row_idx, column=col_idx).value = value
            else:
                source_wb = load_workbook(source_file, data_only=True)
                vda_source = source_wb.active

            def safe_int(value):
                try:
                    return int(value)
                except (TypeError, ValueError):
                    return 0

            values_1 = [[safe_int(vda_source.cell(row=row, column=col).value) for col in range(3, 11)] for row in range(2, 20)]

            values_2 = [[safe_int(vda_source.cell(row=row, column=col).value) for col in range(3, 11)] for row in range(20, 54)]
        except Exception as e:
            return f":( Nepavyko nuskaityti failo: {e}", 500

        # Load a template workbook (assumes 'tuscias.xlsx' is present in the container)
        try:
            template_wb = load_workbook("tuscias_ver.xlsx")
            ataskaita = template_wb.active

            # Įklijuoja pirmą stulpą
            start_row_1, start_col = 35, 13
            for i, row in enumerate(values_1):
                for j, value in enumerate(row):
                    ataskaita.cell(row=start_row_1 + i, column=start_col + j, value=value)

            # Įklijuoja antrą stulpą
            start_row_2 = 60
            for i, row in enumerate(values_2):
                for j, value in enumerate(row):
                    ataskaita.cell(row=start_row_2 + i, column=start_col + j, value=value)
                    
        except Exception as e:
            return f"Error processing template: {e}", 500

        # Save modified workbook into a BytesIO stream and send it for download
        output = io.BytesIO()
        template_wb.save(output)
        output.seek(0)
        return send_file(output,
                         as_attachment=True,
                         download_name="Vertikali_ataskaita.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    return render_template_string(HTML_TEMPLATE)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)