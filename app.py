from flask import Flask, render_template_string, request, send_file
from openpyxl import load_workbook
import io

app = Flask(__name__)

# A very simple HTML form for uploading a file
HTML_TEMPLATE = """
<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>Ataskaitų kepėja Angela</title>
  </head>
  <body>
    <h1>Įkelkite excel failą iš Palantire kad sugeneruoti vertikalią ataskaitą Kultūros ministerijai</h1>
    <form action="/" method="post" enctype="multipart/form-data">
      <input type="file" name="source_file" accept=".xlsx" required>
      <br><br>
      <input type="submit" value="Atsisiųsti ataskaitą">
    </form>
  </body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "source_file" not in request.files:
            return "Failas neįkeltas", 400

        source_file = request.files["source_file"]

        # Load uploaded workbook using openpyxl (assumes source file has data in D2:D8)
        try:
            source_wb = load_workbook(source_file, data_only=True)
            vda_source = source_wb.active
            # Pirmas stulpelis lentelės
            values_1 = [[int(vda_source.cell(row=row, column=col).value) if vda_source.cell(row=row, column=col).value is not None else 0 for col in range(3, 11)] for row in range(2, 20)]

            # antras stulpelis
            values_2 = [[int(vda_source.cell(row=row, column=col).value) if vda_source.cell(row=row, column=col).value is not None else 0 for col in range(3, 11)] for row in range(20, 54)]
        except Exception as e:
            return f"Nepavyko nuskaityti failo: {e}", 500

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