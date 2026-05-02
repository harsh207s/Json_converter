import os
import json
import uuid
import time
import shutil
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file, url_for, after_this_request
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
from datetime import datetime
import re
import threading
import gc
import openpyxl
import csv
from io import StringIO

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['ALLOWED_EXTENSIONS'] = {'pdf', 'xlsx', 'xls', 'csv'}

# Create directories
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Store conversion status
conversion_status = {}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def extract_table_data_fast(text):
    """Fast table extraction from PDF"""
    lines = text.split('\n')
    tables = []
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        
        if ('|' in line or '\t' in line or (line.count(' ') > 2 and len(line) > 30)):
            table_rows = []
            
            while i < len(lines) and lines[i].strip():
                current = lines[i].strip()
                
                if '|' in current:
                    cols = [c.strip() for c in current.split('|') if c.strip()]
                elif '\t' in current:
                    cols = [c.strip() for c in current.split('\t') if c.strip()]
                else:
                    parts = current.split()
                    if len(parts) >= 2:
                        cols = parts
                    else:
                        cols = [current]
                
                if len(cols) > 1:
                    table_rows.append(cols)
                i += 1
                
                if len(table_rows) > 1000:
                    break
            
            if len(table_rows) >= 2:
                tables.append(table_rows)
        else:
            i += 1
    
    return tables

def process_pdf(pdf_path, conversion_id):
    """Process PDF file"""
    start_time = time.time()
    
    conversion_status[conversion_id] = {
        "status": "processing",
        "progress": 0,
        "total_pages": 0,
        "current_page": 0,
        "message": "Processing PDF..."
    }
    
    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        conversion_status[conversion_id]["total_pages"] = total_pages
        
        all_records = []
        all_tables = []
        table_id = 0
        record_id = 0
        
        batch_size = 50
        processed = 0
        
        for batch_start in range(0, total_pages, batch_size):
            batch_end = min(batch_start + batch_size, total_pages)
            batch_records = []
            
            for page_num in range(batch_start, batch_end):
                page = doc[page_num]
                text = page.get_text("text")
                
                if text.strip():
                    tables = extract_table_data_fast(text)
                    
                    for table_data in tables:
                        if len(table_data) >= 2:
                            table_id += 1
                            headers = table_data[0]
                            
                            if len(headers) > 15:
                                headers = headers[:15]
                            
                            for row_data in table_data[1:]:
                                row_dict = {}
                                for i, header in enumerate(headers):
                                    if i < len(row_data):
                                        value = row_data[i]
                                        if value and value.replace('.', '').replace('-', '').isdigit():
                                            try:
                                                value = float(value) if '.' in value else int(value)
                                            except:
                                                pass
                                        row_dict[header] = value
                                    else:
                                        row_dict[header] = ""
                                
                                if row_dict and any(row_dict.values()):
                                    record_id += 1
                                    row_dict["record_id"] = record_id
                                    row_dict["page_number"] = page_num + 1
                                    row_dict["table_id"] = table_id
                                    batch_records.append(row_dict)
                            
                            all_tables.append({
                                "table_id": table_id,
                                "page": page_num + 1,
                                "headers": headers,
                                "rows_count": len(table_data) - 1,
                                "columns_count": len(headers)
                            })
                
                processed += 1
                
                if processed % 5 == 0 or processed == total_pages:
                    progress = (processed / total_pages) * 100
                    elapsed = time.time() - start_time
                    conversion_status[conversion_id].update({
                        "progress": round(progress, 1),
                        "current_page": processed,
                        "message": f"Processing page {processed} of {total_pages}...",
                        "speed": round(processed / elapsed, 2) if elapsed > 0 else 0
                    })
            
            all_records.extend(batch_records)
            gc.collect()
        
        doc.close()
        
        total_time = time.time() - start_time
        
        result = {
            "conversion_info": {
                "tool": "Multi-Format to JSON Converter by Harshu Airy",
                "version": "2.0.0",
                "conversion_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "processing_time_seconds": round(total_time, 2),
                "status": "success",
                "source_type": "PDF"
            },
            "statistics": {
                "total_pages": total_pages,
                "total_tables": table_id,
                "total_records": len(all_records),
                "pages_processed": processed,
                "average_speed": round(len(all_records) / total_time, 2) if total_time > 0 else 0,
                "file_name": os.path.basename(pdf_path)
            },
            "tables": all_tables,
            "records": all_records,
            "summary": {
                "records_by_table": [
                    {
                        "table_id": t["table_id"],
                        "page": t["page"],
                        "rows": t["rows_count"],
                        "columns": t["columns_count"]
                    } for t in all_tables
                ]
            }
        }
        
        return result
        
    except Exception as e:
        conversion_status[conversion_id]["status"] = "error"
        conversion_status[conversion_id]["error"] = str(e)
        return None

def process_excel(excel_path, conversion_id):
    """Process Excel file (XLSX, XLS)"""
    start_time = time.time()
    
    conversion_status[conversion_id] = {
        "status": "processing",
        "progress": 0,
        "message": "Processing Excel file..."
    }
    
    try:
        # Read Excel file
        excel_file = pd.ExcelFile(excel_path)
        sheet_names = excel_file.sheet_names
        total_sheets = len(sheet_names)
        
        all_records = []
        all_tables = []
        record_id = 0
        table_id = 0
        
        for sheet_idx, sheet_name in enumerate(sheet_names):
            progress = (sheet_idx / total_sheets) * 100
            conversion_status[conversion_id].update({
                "progress": round(progress, 1),
                "message": f"Processing sheet {sheet_idx + 1} of {total_sheets}: {sheet_name}",
                "current_sheet": sheet_name
            })
            
            # Read sheet data
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            
            if not df.empty:
                table_id += 1
                headers = df.columns.tolist()
                
                # Convert DataFrame to records
                records = df.to_dict('records')
                
                for record in records:
                    record_id += 1
                    record["record_id"] = record_id
                    record["sheet_name"] = sheet_name
                    record["table_id"] = table_id
                    # Clean NaN values
                    for key, value in list(record.items()):
                        if pd.isna(value):
                            record[key] = None
                        elif isinstance(value, (pd.Timestamp, datetime)):
                            record[key] = value.isoformat()
                    all_records.append(record)
                
                all_tables.append({
                    "table_id": table_id,
                    "sheet_name": sheet_name,
                    "headers": headers,
                    "rows_count": len(df),
                    "columns_count": len(headers)
                })
        
        total_time = time.time() - start_time
        
        result = {
            "conversion_info": {
                "tool": "Multi-Format to JSON Converter by Harshu Airy",
                "version": "2.0.0",
                "conversion_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "processing_time_seconds": round(total_time, 2),
                "status": "success",
                "source_type": "Excel"
            },
            "statistics": {
                "total_sheets": total_sheets,
                "total_tables": table_id,
                "total_records": len(all_records),
                "average_speed": round(len(all_records) / total_time, 2) if total_time > 0 else 0,
                "file_name": os.path.basename(excel_path)
            },
            "tables": all_tables,
            "records": all_records,
            "summary": {
                "sheets_processed": sheet_names,
                "total_rows": len(all_records)
            }
        }
        
        return result
        
    except Exception as e:
        conversion_status[conversion_id]["status"] = "error"
        conversion_status[conversion_id]["error"] = str(e)
        return None

def process_csv(csv_path, conversion_id):
    """Process CSV file"""
    start_time = time.time()
    
    conversion_status[conversion_id] = {
        "status": "processing",
        "progress": 0,
        "message": "Processing CSV file..."
    }
    
    try:
        # Read CSV file
        df = pd.read_csv(csv_path)
        total_rows = len(df)
        
        all_records = []
        record_id = 0
        
        # Convert to records
        records = df.to_dict('records')
        
        for record in records:
            record_id += 1
            record["record_id"] = record_id
            # Clean NaN values
            for key, value in list(record.items()):
                if pd.isna(value):
                    record[key] = None
            all_records.append(record)
            
            # Update progress
            if record_id % 100 == 0:
                progress = (record_id / total_rows) * 100
                conversion_status[conversion_id].update({
                    "progress": round(progress, 1),
                    "message": f"Processing row {record_id} of {total_rows}...",
                    "current_row": record_id
                })
        
        total_time = time.time() - start_time
        
        result = {
            "conversion_info": {
                "tool": "Multi-Format to JSON Converter by Harshu Airy",
                "version": "2.0.0",
                "conversion_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "processing_time_seconds": round(total_time, 2),
                "status": "success",
                "source_type": "CSV"
            },
            "statistics": {
                "total_rows": total_rows,
                "total_records": len(all_records),
                "average_speed": round(len(all_records) / total_time, 2) if total_time > 0 else 0,
                "file_name": os.path.basename(csv_path)
            },
            "records": all_records,
            "summary": {
                "total_rows": len(all_records),
                "columns": list(df.columns)
            }
        }
        
        return result
        
    except Exception as e:
        conversion_status[conversion_id]["status"] = "error"
        conversion_status[conversion_id]["error"] = str(e)
        return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': 'Please upload a valid file (PDF, Excel, or CSV)'}), 400
        
        conversion_id = str(uuid.uuid4())[:8]
        
        original_filename = secure_filename(file.filename)
        file_extension = original_filename.rsplit('.', 1)[1].lower()
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{conversion_id}_{original_filename}")
        file.save(temp_path)
        
        file_size = os.path.getsize(temp_path) / (1024 * 1024)
        
        def process():
            result = None
            if file_extension == 'pdf':
                result = process_pdf(temp_path, conversion_id)
            elif file_extension in ['xlsx', 'xls']:
                result = process_excel(temp_path, conversion_id)
            elif file_extension == 'csv':
                result = process_csv(temp_path, conversion_id)
            
            if result:
                json_filename = f"{conversion_id}_output.json"
                json_path = os.path.join(app.config['OUTPUT_FOLDER'], json_filename)
                
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(result, f, indent=2, ensure_ascii=False)
                
                conversion_status[conversion_id]["json_file"] = json_filename
                conversion_status[conversion_id]["result"] = result
                conversion_status[conversion_id]["status"] = "completed"
                conversion_status[conversion_id]["progress"] = 100
                
                try:
                    os.remove(temp_path)
                except:
                    pass
            else:
                conversion_status[conversion_id]["status"] = "error"
                conversion_status[conversion_id]["error"] = "Conversion failed"
        
        thread = threading.Thread(target=process)
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'success': True,
            'conversion_id': conversion_id,
            'file_size': round(file_size, 2),
            'file_type': file_extension,
            'message': 'Conversion started successfully'
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/status/<conversion_id>')
def get_status(conversion_id):
    if conversion_id in conversion_status:
        status = conversion_status[conversion_id]
        
        if status["status"] == "completed":
            return jsonify({
                "status": "completed",
                "progress": 100,
                "result": status.get("result"),
                "json_file": status.get("json_file")
            })
        elif status["status"] == "error":
            return jsonify({
                "status": "error",
                "error": status.get("error")
            })
        else:
            return jsonify({
                "status": "processing",
                "progress": status.get("progress", 0),
                "message": status.get("message", "Processing..."),
                "current_page": status.get("current_page", 0),
                "total_pages": status.get("total_pages", 0),
                "speed": status.get("speed", 0)
            })
    
    return jsonify({"status": "not_found"}), 404

@app.route('/download/<conversion_id>')
def download_file(conversion_id):
    try:
        if conversion_id in conversion_status:
            status = conversion_status[conversion_id]
            
            if status.get("status") == "completed" and status.get("json_file"):
                json_filename = status["json_file"]
                json_path = os.path.join(app.config['OUTPUT_FOLDER'], json_filename)
                
                if os.path.exists(json_path):
                    result = status.get("result", {})
                    original_name = result.get("statistics", {}).get("file_name", "converted")
                    base_name = os.path.splitext(original_name)[0]
                    download_name = f"{base_name}_converted.json"
                    
                    return send_file(
                        json_path,
                        as_attachment=True,
                        download_name=download_name,
                        mimetype='application/json'
                    )
        
        return jsonify({'error': 'File not found'}), 404
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/preview/<conversion_id>')
def preview_data(conversion_id):
    try:
        if conversion_id in conversion_status:
            status = conversion_status[conversion_id]
            
            if status.get("status") == "completed" and status.get("result"):
                result = status["result"]
                
                preview = {
                    "conversion_info": result.get("conversion_info", {}),
                    "statistics": result.get("statistics", {}),
                    "sample_records": result.get("records", [])[:10],
                    "summary": result.get("summary", {})
                }
                
                return jsonify(preview)
        
        return jsonify({'error': 'No data available'}), 404
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/cleanup', methods=['POST'])
def cleanup_files():
    try:
        current_time = time.time()
        
        for folder in [app.config['OUTPUT_FOLDER'], app.config['UPLOAD_FOLDER']]:
            for filename in os.listdir(folder):
                filepath = os.path.join(folder, filename)
                if os.path.isfile(filepath):
                    if current_time - os.path.getctime(filepath) > 3600:
                        os.remove(filepath)
        
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

if __name__ == '__main__':
    print("\n" + "="*60)
    print("🚀 Multi-Format to JSON Converter by Harshu Airy")
    print("="*60)
    print(f"📍 Server running at: http://localhost:5000")
    print(f"📁 Supported formats: PDF, Excel (XLSX/XLS), CSV")
    print("="*60 + "\n")
    app.run(debug=False, host='0.0.0.0', port=5000, threaded=True)