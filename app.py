from flask import Flask, request, jsonify, send_file
import requests
import pandas as pd
import time
from datetime import datetime
from io import BytesIO

app = Flask(__name__)

API_URL = "https://s4c-lms-api.onrender.com/api/attendance/daily"
HEADERS = {
    "Content-Type": "application/json",
    "X-API-Key": "mKhFtHShfAIwo7YV6zfpckXriHsQsgO2"  # Replace with your API key
}

def fetch_attendance(org_emp_code, attendance_date, batch_id):
    params = {
        "attendanceDate": attendance_date,
        "batch_id": batch_id,
        "org_emp_code": org_emp_code
    }
    try:
        response = requests.get(API_URL, headers=HEADERS, params=params)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}

@app.route('/fetch-attendance', methods=['POST'])
def fetch_attendance_endpoint():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400
    
    file = request.files['file']
    attendance_date = request.form.get("attendance_date")
    batch_id = request.form.get("batch_id")
    
    if not attendance_date or not batch_id:
        return jsonify({"error": "Missing required parameters: attendance_date or batch_id"}), 400
    
    try:
        date_obj = datetime.strptime(attendance_date, "%Y-%m-%d")
        filename = f"S4Carlisle_{date_obj.day:02d}_{date_obj.month:02d}_{str(date_obj.year)[-2:]}.xlsx"
    except ValueError:
        return jsonify({"error": "Invalid attendance_date format. Use YYYY-MM-DD."}), 400
    
    try:
        org_emp_codes = [line.strip() for line in file.stream.read().decode('utf-8').splitlines()]
    except Exception as e:
        return jsonify({"error": f"Error reading file: {str(e)}"}), 400

    results = []

    for code in org_emp_codes:
        result = fetch_attendance(code, attendance_date, batch_id)
        if "Successattendance" in result and isinstance(result["Successattendance"], list):
            for record in result["Successattendance"]:
                results.append({
                    "org_emp_code": record.get("org_emp_code"),
                    "attendanceDate": attendance_date,
                    "in_time": record.get("in_time"),
                    "out_time": record.get("out_time")
                })
        else:
            results.append({
                "org_emp_code": code,
                "attendanceDate": attendance_date,
                "in_time": None,
                "out_time": None
            })
        time.sleep(0.2)

    # Create the Excel file in memory
    df = pd.DataFrame(results)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    # Send the file as a response
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
