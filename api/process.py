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
                tmp_file_path = tmp_file.name

            try:
                # Process the file
                success, message, result_df = self.process_uploaded_file(tmp_file_path, process_type)
                
                if success and result_df is not None:
                    # Convert DataFrame to JSON-serializable format
                    data = self.prepare_response_data(result_df)
                    
                    response = {
                        'success': True,
                        'message': message,
                        'data': data
                    }
                else:
                    response = {
                        'success': False,
                        'message': message,
                        'data': None
                    }

                self.send_json_response(200, response)

            finally:
                # Clean up temporary file
                if os.path.exists(tmp_file_path):
                    os.unlink(tmp_file_path)

        except Exception as e:
            logging.error(f"Error processing request: {str(e)}", exc_info=True)
            self.send_error_response(500, f"Internal server error: {str(e)}")

    def process_uploaded_file(self, file_path, process_type):
        """Process the uploaded file using the appropriate converter"""
        try:
            # Process file based on type
            if process_type == "Windows":
                df, _ = convertWindow.process_file(file_path)
            elif process_type == "Door":
                df, _ = convertDoor.process_file(file_path)
            else:
                return False, "Unknown process type", None

            # Process cutting data
            success, message, result_df = self.process_cutting_data(df)
            return success, message, result_df

        except Exception as e:
            error_message = f"Error processing file: {str(e)}"
            logging.error(error_message, exc_info=True)
            return False, error_message, None

    def process_cutting_data(self, df):
        """Process cutting data - core optimization logic"""
        try:
            logging.info("Starting data processing")
            
            # Check required columns
            required_columns = ['Material Name', 'Qty', 'Length', 'Order No', 'Bin No']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
            
            # Create temporary DataFrame for sorting and calculations
            temp_df = df.copy()
            temp_df['original_index'] = temp_df.index
            temp_df = temp_df.sort_values(['Material Name', 'Qty', 'Length', 'Order No', 'Bin No'], 
                                          ascending=[True, True, False, True, True])
            
            cutting_info = defaultdict(list)
            material_total_lengths = defaultdict(float)
            material_max_cutting_id = defaultdict(int)
            processed_rows = set()

            for (material, qty), material_group in temp_df.groupby(['Material Name', 'Qty']):
                material_length = get_material_length(material)
                logging.info(f"Processing material {material}, qty {qty}, standard length: {material_length}")
                
                all_lengths = material_group['Length'].tolist()
                cutting_id = material_max_cutting_id[material] + 1
                
                while all_lengths:
                    best_combination = self.find_best_combination(all_lengths, material_length - 6, 4, 10)
                    remaining = material_length - sum(best_combination) - (len(best_combination) - 1) * 4 - 6
                    logging.debug(f"Cutting ID {cutting_id} best combination: {best_combination}, remaining: {remaining}")
                    
                    # If no valid combination found, process single lengths
                    if not best_combination:
                        if all_lengths:
                            single_length = all_lengths[0]
                            logging.warning(f"No optimal combination found, processing single length: {single_length}")
                            
                            # Find corresponding row
                            unprocessed_rows = material_group[
                                (material_group['Length'] == single_length) & 
                                (~material_group.index.isin(processed_rows))
                            ]
                            
                            if not unprocessed_rows.empty:
                                row = unprocessed_rows.iloc[0]
                                key = (row['Material Name'], row['Qty'], row['Length'], row['Order No'], row['Bin No'])
                                cutting_info[(material, qty)].append((key, cutting_id, 1, row['original_index']))
                                
                                processed_rows.add(row.name)
                                all_lengths.remove(single_length)
                                material_total_lengths[(material, qty)] += single_length
                            else:
                                all_lengths.remove(single_length)
                                logging.warning(f"No unprocessed row found for length {single_length}, removing")
                        
                        cutting_id += 1
                        continue
                    
                    for pieces_id, length in enumerate(best_combination, 1):
                        unprocessed_rows = material_group[
                            (material_group['Length'] == length) & 
                            (~material_group.index.isin(processed_rows))
                        ]
                        
                        if not unprocessed_rows.empty:
                            row = unprocessed_rows.iloc[0]
                            key = (row['Material Name'], row['Qty'], row['Length'], row['Order No'], row['Bin No'])
                            cutting_info[(material, qty)].append((key, cutting_id, pieces_id, row['original_index']))
                            
                            processed_rows.add(row.name)
                            all_lengths.remove(length)
                            material_total_lengths[(material, qty)] += length
                        else:
                            logging.warning(f"No unprocessed row found for length {length}, skipping")
                    
                    cutting_id += 1
                
                material_max_cutting_id[material] = cutting_id - 1
                logging.info(f"Material {material}, qty {qty} processing complete, total length: {material_total_lengths[(material, qty)]:.2f}")

            logging.info("Cutting information calculation complete")

            # Create result DataFrame maintaining original order
            result_df = df.copy()
            result_df['Cutting ID'] = 0
            result_df['Pieces ID'] = 0

            # Fill Cutting ID and Pieces ID
            for (material, qty), info_list in cutting_info.items():
                for (key, cutting_id, pieces_id, original_index) in info_list:
                    result_df.loc[original_index, 'Cutting ID'] = cutting_id
                    result_df.loc[original_index, 'Pieces ID'] = pieces_id

            logging.info("Cutting ID and Pieces ID filling complete")
            
            return True, "Data processing successful", result_df
        
        except Exception as e:
            logging.error(f"Error processing data: {str(e)}", exc_info=True)
            return False, f"Error processing data: {str(e)}", None

    def find_best_combination(self, lengths, target_length, cut_loss=4, min_remaining=10):
        """Find the best combination of lengths that fit within target_length"""
        lengths = sorted(lengths, reverse=True)
        best_combination = []
        best_remaining = target_length
        current_combination = []
        current_length = 0

        def backtrack(index):
            nonlocal best_combination, best_remaining, current_combination, current_length

            # If current combination is better than best, update best
            if len(current_combination) > 0 and target_length - current_length < best_remaining:
                best_combination = current_combination.copy()
                best_remaining = target_length - current_length

            # Base case: considered all lengths or remaining length too small
            if index == len(lengths) or target_length - current_length < min_remaining:
                return

            # Try adding current length
            if current_length + lengths[index] + cut_loss <= target_length:
                current_combination.append(lengths[index])
                current_length += lengths[index] + cut_loss
                backtrack(index + 1)
                current_combination.pop()
                current_length -= lengths[index] + cut_loss

            # Try not adding current length
            backtrack(index + 1)

        backtrack(0)
        return best_combination

    def prepare_response_data(self, df):
        """Prepare DataFrame for JSON response"""
        # Convert DataFrame to dict
        data = {
            'rows': df.to_dict('records'),
            'columns': df.columns.tolist(),
            'totalLength': float(df['Length'].sum()) if 'Length' in df.columns else 0,
            'materialTypes': int(df['Material Name'].nunique()) if 'Material Name' in df.columns else 0,
            'maxCuttingId': int(df['Cutting ID'].max()) if 'Cutting ID' in df.columns else 0
        }
        
        # Convert numpy types to Python types for JSON serialization
        for row in data['rows']:
            for key, value in row.items():
                if pd.isna(value):
                    row[key] = None
                elif hasattr(value, 'item'):  # numpy types
                    row[key] = value.item()
        
        return data

    def send_json_response(self, status_code, data):
        """Send JSON response"""
        self.send_response(status_code)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()
        
        response_json = json.dumps(data, ensure_ascii=False, indent=2)
        self.wfile.write(response_json.encode('utf-8'))

    def send_error_response(self, status_code, message):
        """Send error response"""
        error_data = {
            'success': False,
            'message': message,
            'data': None
        }
        self.send_json_response(status_code, error_data)

    def do_OPTIONS(self):
        """Handle CORS preflight requests"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

# Configure logging
logging.basicConfig(level=logging.INFO)