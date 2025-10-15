from http.server import BaseHTTPRequestHandler
import json
import tempfile
import os
import sys
from collections import defaultdict
import logging
import traceback
from io import BytesIO
import cgi
import openpyxl
import pandas as pd
import numpy as np

# Add the parent directory to Python path for imports
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

# Import modules at the top level to avoid import issues in Vercel
try:
    # Try importing from parent directory first
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
    import convertWindow
    import convertDoor
    from app import process_cutting_data
except ImportError as e:
    logging.error(f"Import error: {e}")
    # Fallback imports for Vercel environment
    try:
        sys.path.append('/var/task')
        sys.path.append('/var/runtime')
        sys.path.append(os.path.dirname(os.path.dirname(__file__)))
        import convertWindow
        import convertDoor
        from app import process_cutting_data
    except ImportError as e2:
        logging.error(f"Fallback import error: {e2}")
        # If imports still fail, we need to handle this gracefully
        convertWindow = None
        convertDoor = None
        process_cutting_data = None

class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        """Handle CORS preflight requests"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_POST(self):
        try:
            # Check if required modules are available
            if convertWindow is None or convertDoor is None or process_cutting_data is None:
                self.send_error_response(500, "Required modules not available. Import error occurred.")
                return
                
            # Set CORS headers
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            self.send_header('Content-Type', 'application/json')
            self.end_headers()

            # Parse the multipart form data
            content_type = self.headers.get('Content-Type', '')
            if not content_type.startswith('multipart/form-data'):
                self.send_error_response(400, "Content-Type must be multipart/form-data")
                return

            # Parse form data
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={'REQUEST_METHOD': 'POST'}
            )

            # Get file and process type
            if 'file' not in form:
                self.send_error_response(400, "No file uploaded")
                return

            file_item = form['file']
            if not file_item.filename:
                self.send_error_response(400, "No file selected")
                return

            process_type = form.getvalue('processType', 'Windows')

            # Read file content directly into memory
            file_content = file_item.file.read()
            
            # Save to temporary file for processing
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_item.filename)[1]) as tmp_file:
                tmp_file.write(file_content)
                tmp_file_path = tmp_file.name

            try:
                # Process the file based on type
                if process_type == 'Windows':
                    df, _ = convertWindow.process_file(tmp_file_path)
                else:  # Door
                    df, _ = convertDoor.process_file(tmp_file_path)
                
                # Process cutting data
                success, message, result_df = process_cutting_data(df)
                
                if not success:
                    self.send_error_response(400, message)
                    return
                
                # Convert DataFrame to the expected format with JSON serialization fix
                def convert_numpy_types(obj):
                    if isinstance(obj, np.integer):
                        return int(obj)
                    elif isinstance(obj, np.floating):
                        return float(obj)
                    elif isinstance(obj, np.ndarray):
                        return obj.tolist()
                    elif pd.isna(obj):  # Handle NaN values
                        return None
                    return obj
                
                rows = result_df.to_dict('records')
                # Convert numpy types and handle NaN values in each row
                for row in rows:
                    for key, value in row.items():
                        row[key] = convert_numpy_types(value)
                
                columns = list(result_df.columns)
                
                # Calculate stats with proper type conversion
                max_cutting_id = result_df['Cutting ID'].max() if 'Cutting ID' in result_df.columns else 0
                stats = {
                    'total_pieces': len(rows),
                    'total_cuts': int(max_cutting_id) if max_cutting_id is not None and not pd.isna(max_cutting_id) else 0,
                    'material_usage': {}
                }
                
                response_data = {
                    'success': True,
                    'data': {
                        'rows': rows,
                        'columns': columns,
                        'stats': stats
                    },
                    'filename': file_item.filename
                }
                
                self.wfile.write(json.dumps(response_data).encode('utf-8'))
                
            finally:
                # Clean up temporary file
                if os.path.exists(tmp_file_path):
                    os.unlink(tmp_file_path)

        except Exception as e:
            logging.error(f"Error processing file: {str(e)}")
            logging.error(traceback.format_exc())
            self.send_error_response(500, f"Internal server error: {str(e)}")

    def send_error_response(self, status_code, message):
        """Send error response with proper headers"""
        self.send_response(status_code)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        
        error_response = {
            'success': False,
            'message': message
        }
        self.wfile.write(json.dumps(error_response).encode('utf-8'))