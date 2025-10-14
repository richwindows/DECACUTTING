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

# Import the processing modules
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from settings import get_material_length
import convertWindow
import convertDoor

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

            # Save uploaded file to temporary location
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_item.filename)[1]) as tmp_file:
                tmp_file.write(file_item.file.read())
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