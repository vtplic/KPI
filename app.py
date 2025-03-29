from flask import Flask, request, send_file, render_template_string
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os
import matplotlib.pyplot as plt
from matplotlib import rcParams

rcParams["font.family"] = "Times New Roman"

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

HTML_TEMPLATE = '''
<!doctype html>
<html>
<head>
    <title>KPI Liên Chiểu</title>
</head>
<body>
    <h2>KPI Bưu cục Liên Chiểu</h2>
    <form action="/process" method="post" enctype="multipart/form-data">
        <input type="file" name="file">
        <br><br>
        <label for="cod_threshold">Tỷ lệ COD tối thiểu:</label>
        <input type="number" name="cod_threshold" step="0.1" value="95">
        <br><br>
        <label for="time_threshold">Tỷ lệ đúng giờ tối thiểu:</label>
        <input type="number" name="time_threshold" step="0.1" value="85">
        <br><br>
        <input type="submit" value="Upload and Process">
    </form>
    <br>
    {% if img_path %}
        <h2>KPI luỹ kế tháng Bưu cục Liên Chiểu</h2>
        <img src="{{ img_path }}" alt="KPI Warning" style="max-width: 30%;">
        <br><br>
        <form action="/download" method="get">
            <button type="submit">Xuất file Excel</button>
        </form>
    {% endif %}
</body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_form():
    return render_template_string(HTML_TEMPLATE)

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    
    cod_threshold = float(request.form.get('cod_threshold', 95))
    time_threshold = float(request.form.get('time_threshold', 85))
    
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()
    
    expected_columns = ["Tuyến", "Phát thành công COD", "Phát thành công đúng giờ"]
    actual_columns = df.columns.tolist()
    
    if not all(col in actual_columns for col in expected_columns):
        return f"Error: Expected columns not found! Columns in file: {actual_columns}"
    
    df_filtered = df[["Tuyến", "Phát thành công COD", "Phát thành công đúng giờ"]].copy()
    df_filtered.columns = ["Nhân viên", "Tỷ lệ COD", "Tỷ lệ đúng giờ"]
    df_filtered["Tỷ lệ COD"] = pd.to_numeric(df_filtered["Tỷ lệ COD"], errors="coerce")
    df_filtered["Tỷ lệ đúng giờ"] = pd.to_numeric(df_filtered["Tỷ lệ đúng giờ"], errors="coerce")
    
    df_filtered = df_filtered[(df_filtered["Tỷ lệ COD"] >= 70) & (df_filtered["Tỷ lệ COD"] < 100)]
    df_filtered = df_filtered.sort_values(by="Tỷ lệ COD", ascending=False)
    
    output_path = os.path.join(RESULT_FOLDER, "KPI_Canh_Bao.xlsx")
    df_filtered.to_excel(output_path, index=False)
    
    wb = load_workbook(output_path)
    ws = wb.active
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    
    for row in range(2, ws.max_row + 1):
        cod_cell = ws[f"B{row}"]
        time_cell = ws[f"C{row}"]
        if cod_cell.value is not None and cod_cell.value < cod_threshold:
            cod_cell.fill = red_fill
        if time_cell.value is not None and time_cell.value < time_threshold:
            time_cell.fill = red_fill
    
    wb.save(output_path)
    
    img_path = os.path.join(RESULT_FOLDER, "KPI_Canh_Bao.png")
    fig, ax = plt.subplots(figsize=(4, len(df_filtered) * 0.2))
    ax.axis('tight')
    ax.axis('off')
    
    df_filtered["Tỷ lệ COD"] = df_filtered["Tỷ lệ COD"].apply(lambda x: f"{x:.2f}%")
    df_filtered["Tỷ lệ đúng giờ"] = df_filtered["Tỷ lệ đúng giờ"].apply(lambda x: f"{x:.2f}%")
    
    table_data = df_filtered.values.tolist()
    table = ax.table(cellText=table_data, colLabels=df_filtered.columns, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.auto_set_column_width([0, 1, 2])
    
    for key, cell in table.get_celld().items():
        cell.set_fontsize(10)
        cell.set_height(0.1)
        if key[0] == 0:
            cell.set_text_props(weight='bold', color='white')
            cell.set_facecolor('#4C72B0')
        if key[0] > 0 and key[1] in [1, 2]:
            value = df_filtered.iloc[key[0] - 1, key[1]]
            if (key[1] == 1 and float(value[:-1]) < cod_threshold) or (key[1] == 2 and float(value[:-1]) < time_threshold):
                cell.set_facecolor("red")
                cell.set_text_props(color='white')
    
    plt.savefig(img_path, bbox_inches='tight', dpi=300)
    
    return render_template_string(HTML_TEMPLATE, img_path="/image")

@app.route('/download')
def download_file():
    output_path = os.path.join(RESULT_FOLDER, "KPI_Canh_Bao.xlsx")
    return send_file(output_path, as_attachment=True)

@app.route('/image')
def get_image():
    img_path = os.path.join(RESULT_FOLDER, "KPI_Canh_Bao.png")
    if os.path.exists(img_path):
        return send_file(img_path, mimetype='image/png')
    return "No image available"

if __name__ == '__main__':
    app.run(debug=True)
