"""report_generator.py
Generate a human-readable HTML report using a simple jinja2 template.
"""
import json
import sys
from jinja2 import Template

HTML_TEMPLATE = '''
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Excel Analyzer Report</title>
<style>
body{font-family: Arial, sans-serif;padding:18px}
h1{font-size:22px}
.section{margin-bottom:18px;padding:12px;border:1px solid #ddd;border-radius:8px}
.table{width:100%;border-collapse:collapse}
.table th,.table td{border:1px solid #ccc;padding:6px;text-align:left}
</style>
</head>
<body>
<h1>Excel Analyzer Report</h1>
<div class="section">
<strong>File:</strong> {{ path }}<br>
<strong>Sheets:</strong> {{ sheet_count }}<br>
<strong>Zip entries:</strong> {{ zip_entry_count }}<br>
<strong>Media files:</strong> {{ media_count }}<br>
</div>


<div class="section">
<h2>Summary</h2>
<ul>
<li>Total cells scanned (estimate): {{ total_cells_scanned_estimate }}</li>
<li>Total formulas: {{ total_formulas }}</li>
<li>Total volatile formulas: {{ total_volatile_formulas }}</li>
<li>Total merged cells: {{ total_merged_cells }}</li>
</ul>
</div>


<div class="section">
<h2>Sheets</h2>
<table class="table">
<thead><tr><th>Sheet</th><th>Rows</th><th>Cols</th><th>Formulas</th><th>Volatile</th><th>Merged</th><th>Hidden rows</th><th>Hidden cols</th><th>Unique styles</th></tr></thead>
<tbody>
{% for name,info in sheets.items() %}
<tr>
<td>{{ name }}</td>
<td>{{ info.max_row }}</td>
<td>{{ info.max_column }}</td>
<td>{{ info.formulas }}</td>
<td>{{ info.volatile_formulas }}</td>
<td>{{ info.merged_cells }}</td>
<td>{{ info.hidden_rows }}</td>
<td>{{ info.hidden_columns }}</td>
<td>{{ info.unique_styles }}</td>
</tr>
{% endfor %}
</tbody>
</table>
</div>


</body>
</html>
'''

def generate_report(analyzer_result: dict, out_path: str):
    tmpl = Template(HTML_TEMPLATE)
    html = tmpl.render(**analyzer_result)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python report_generator.py <analyzer_json_path> <out_report.html>')
        sys.exit(1)
    import json
    ajson = sys.argv[1]
    out = sys.argv[2]
    with open(ajson, 'r', encoding='utf-8') as f:
        data = json.load(f)
    generate_report(data, out)
    print('Report written to', out)