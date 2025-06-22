from flask import Flask, request, jsonify, render_template
import openpyxl
import os
import sys

BASE_DIR = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
app = Flask(__name__, template_folder=os.path.join(BASE_DIR, 'templates'))
EXCEL_FILE = os.path.join(os.getcwd(), 'data.xlsx')


def load_workbook():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["ID", "Name"])
        wb.save(EXCEL_FILE)
    return openpyxl.load_workbook(EXCEL_FILE)


def read_data():
    wb = load_workbook()
    ws = wb.active
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        data.append({"id": row[0], "name": row[1]})
    return data


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/data')
def list_data():
    return jsonify(read_data())


@app.route('/api/add', methods=['POST'])
def add_row():
    id_ = request.json.get('id')
    name = request.json.get('name')
    if not id_ or not name:
        return jsonify({"error": "id and name required"}), 400
    wb = load_workbook()
    ws = wb.active
    ws.append([id_, name])
    wb.save(EXCEL_FILE)
    return jsonify({"success": True})


@app.route('/api/update', methods=['POST'])
def update_row():
    id_ = request.json.get('id')
    name = request.json.get('name')
    if not id_ or not name:
        return jsonify({"error": "id and name required"}), 400
    wb = load_workbook()
    ws = wb.active
    for row in ws.iter_rows(min_row=2):
        if str(row[0].value) == str(id_):
            row[1].value = name
            wb.save(EXCEL_FILE)
            return jsonify({"success": True})
    return jsonify({"error": "ID not found"}), 404


@app.route('/api/delete', methods=['POST'])
def delete_row():
    id_ = request.json.get('id')
    if not id_:
        return jsonify({"error": "id required"}), 400
    wb = load_workbook()
    ws = wb.active
    row_to_delete = None
    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        if str(row[0].value) == str(id_):
            row_to_delete = idx
            break
    if row_to_delete:
        ws.delete_rows(row_to_delete)
        wb.save(EXCEL_FILE)
        return jsonify({"success": True})
    else:
        return jsonify({"error": "ID not found"}), 404


if __name__ == '__main__':
    app.run(debug=True)
