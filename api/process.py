from http.server import BaseHTTPRequestHandler
import json
import tempfile
import os
import pandas as pd
from collections import defaultdict
import logging
import traceback
from io import BytesIO
import cgi
import openpyxl

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
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
            
            # Process the file based on type
            if process_type == 'Windows':
                processed_data = self.process_window_file(file_content, file_item.filename)
            else:  # Doors
                processed_data = self.process_door_file(file_content, file_item.filename)

            # Apply cutting optimization
            optimized_data = self.process_cutting_data(processed_data)

            # Prepare response
            response_data = {
                'success': True,
                'data': {
                    'rows': optimized_data.to_dict('records'),
                    'columns': list(optimized_data.columns),
                    'stats': {
                        'total_pieces': len(optimized_data),
                        'total_cuts': len(optimized_data['Cutting ID'].unique()) if 'Cutting ID' in optimized_data.columns else 0,
                        'material_usage': self.calculate_material_usage(optimized_data)
                    }
                },
                'filename': file_item.filename
            }

            self.send_json_response(response_data)

        except Exception as e:
            logging.error(f"Error processing file: {str(e)}")
            logging.error(traceback.format_exc())
            self.send_error_response(500, f"Internal server error: {str(e)}")

    def process_window_file(self, file_content, filename):
        """Process window file content"""
        try:
            # Load Excel file from memory
            workbook = openpyxl.load_workbook(BytesIO(file_content), data_only=True)
            
            # Get material color suffix function
            def get_material_color_suffix(color_value):
                if pd.isna(color_value) or color_value == '':
                    return ''
                color_str = str(color_value).strip().upper()
                if color_str in ['WHITE', 'BLANC']:
                    return 'WH'
                elif color_str in ['BROWN', 'BRUN']:
                    return 'BR'
                elif color_str in ['BLACK', 'NOIR']:
                    return 'BL'
                else:
                    return color_str[:2] if len(color_str) >= 2 else color_str
            
            # Process the workbook
            data_rows = []
            
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # Find data starting row (skip headers)
                start_row = 2  # Assuming headers are in row 1
                
                for row in sheet.iter_rows(min_row=start_row, values_only=True):
                    if not any(row):  # Skip empty rows
                        continue
                        
                    # Extract relevant data (adjust column indices based on your Excel structure)
                    try:
                        width = row[0] if row[0] is not None else 0
                        height = row[1] if row[1] is not None else 0
                        color = row[2] if len(row) > 2 else ''
                        quantity = row[3] if len(row) > 3 and row[3] is not None else 1
                        
                        if width > 0 and height > 0:
                            color_suffix = get_material_color_suffix(color)
                            material_name = f"Window_{color_suffix}" if color_suffix else "Window"
                            
                            data_rows.append({
                                'Width': width,
                                'Height': height,
                                'Color': color,
                                'Quantity': quantity,
                                'Material': material_name,
                                'Type': 'Window'
                            })
                    except (IndexError, TypeError, ValueError):
                        continue
            
            return pd.DataFrame(data_rows)
            
        except Exception as e:
            raise Exception(f"Error processing window file: {str(e)}")

    def process_door_file(self, file_content, filename):
        """Process door file content"""
        try:
            # Load Excel file from memory
            workbook = openpyxl.load_workbook(BytesIO(file_content), data_only=True)
            
            # Get material color suffix function
            def get_material_color_suffix(color_value):
                if pd.isna(color_value) or color_value == '':
                    return ''
                color_str = str(color_value).strip().upper()
                if color_str in ['WHITE', 'BLANC']:
                    return 'WH'
                elif color_str in ['BROWN', 'BRUN']:
                    return 'BR'
                elif color_str in ['BLACK', 'NOIR']:
                    return 'BL'
                else:
                    return color_str[:2] if len(color_str) >= 2 else color_str
            
            # Process the workbook
            data_rows = []
            
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                # Find data starting row (skip headers)
                start_row = 2  # Assuming headers are in row 1
                
                for row in sheet.iter_rows(min_row=start_row, values_only=True):
                    if not any(row):  # Skip empty rows
                        continue
                        
                    # Extract relevant data (adjust column indices based on your Excel structure)
                    try:
                        width = row[0] if row[0] is not None else 0
                        height = row[1] if row[1] is not None else 0
                        color = row[2] if len(row) > 2 else ''
                        quantity = row[3] if len(row) > 3 and row[3] is not None else 1
                        
                        if width > 0 and height > 0:
                            color_suffix = get_material_color_suffix(color)
                            material_name = f"Door_{color_suffix}" if color_suffix else "Door"
                            
                            data_rows.append({
                                'Width': width,
                                'Height': height,
                                'Color': color,
                                'Quantity': quantity,
                                'Material': material_name,
                                'Type': 'Door'
                            })
                    except (IndexError, TypeError, ValueError):
                        continue
            
            return pd.DataFrame(data_rows)
            
        except Exception as e:
            raise Exception(f"Error processing door file: {str(e)}")

    def get_material_length(self, material_name):
        """Get material length - simplified version"""
        # Default material lengths
        default_lengths = {
            'Window_WH': 6000,
            'Window_BR': 6000,
            'Window_BL': 6000,
            'Door_WH': 6000,
            'Door_BR': 6000,
            'Door_BL': 6000,
        }
        return default_lengths.get(material_name, 6000)

    def process_cutting_data(self, df):
        """Process cutting data with optimization"""
        if df.empty:
            return df
            
        # Add cutting optimization logic
        df_expanded = []
        
        for _, row in df.iterrows():
            quantity = int(row.get('Quantity', 1))
            for i in range(quantity):
                new_row = row.copy()
                new_row['Piece_ID'] = f"{row.get('Type', 'Item')}_{len(df_expanded) + 1}"
                df_expanded.append(new_row)
        
        df_expanded = pd.DataFrame(df_expanded)
        
        # Apply cutting optimization
        optimized_df = self.find_best_combination(df_expanded)
        
        return optimized_df

    def find_best_combination(self, df):
        """Find best cutting combination"""
        if df.empty:
            return df
            
        # Group by material type
        material_groups = df.groupby('Material')
        result_rows = []
        cutting_id = 1
        
        for material, group in material_groups:
            material_length = self.get_material_length(material)
            pieces = group.to_dict('records')
            
            # Simple bin packing algorithm
            current_cut = []
            current_length = 0
            
            for piece in pieces:
                piece_length = max(piece.get('Width', 0), piece.get('Height', 0))
                
                if current_length + piece_length <= material_length:
                    current_cut.append(piece)
                    current_length += piece_length
                else:
                    # Finalize current cut
                    if current_cut:
                        for i, cut_piece in enumerate(current_cut):
                            cut_piece['Cutting ID'] = f"CUT_{cutting_id}"
                            cut_piece['Position'] = i + 1
                            cut_piece['Material_Length'] = material_length
                            cut_piece['Used_Length'] = current_length
                            result_rows.append(cut_piece)
                        cutting_id += 1
                    
                    # Start new cut
                    current_cut = [piece]
                    current_length = piece_length
            
            # Handle remaining pieces
            if current_cut:
                for i, cut_piece in enumerate(current_cut):
                    cut_piece['Cutting ID'] = f"CUT_{cutting_id}"
                    cut_piece['Position'] = i + 1
                    cut_piece['Material_Length'] = material_length
                    cut_piece['Used_Length'] = current_length
                    result_rows.append(cut_piece)
                cutting_id += 1
        
        return pd.DataFrame(result_rows)

    def calculate_material_usage(self, df):
        """Calculate material usage statistics"""
        if df.empty or 'Used_Length' not in df.columns or 'Material_Length' not in df.columns:
            return 0
            
        total_used = df['Used_Length'].sum()
        total_available = df['Material_Length'].sum()
        
        return round((total_used / total_available * 100), 2) if total_available > 0 else 0

    def send_json_response(self, data):
        """Send JSON response"""
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        
        response_json = json.dumps(data, ensure_ascii=False, default=str)
        self.wfile.write(response_json.encode('utf-8'))

    def send_error_response(self, status_code, message):
        """Send error response"""
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        
        error_data = {
            'success': False,
            'message': message
        }
        
        response_json = json.dumps(error_data, ensure_ascii=False)
        self.wfile.write(response_json.encode('utf-8'))

    def do_OPTIONS(self):
        """Handle CORS preflight requests"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()