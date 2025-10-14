from http.server import BaseHTTPRequestHandler
import json
import pandas as pd
from io import BytesIO
import os

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
            # Set CORS headers
            self.send_response(200)
            self.send_header('Access-Control-Allow-Origin', '*')
            
            # Read request body
            content_length = int(self.headers.get('Content-Length', 0))
            post_data = self.rfile.read(content_length)
            
            # Parse JSON data
            try:
                request_data = json.loads(post_data.decode('utf-8'))
            except json.JSONDecodeError:
                self.send_error_response(400, "Invalid JSON data")
                return

            # Extract parameters
            data = request_data.get('data')
            file_format = request_data.get('format', 'excel')
            original_filename = request_data.get('filename', 'processed_data')

            if not data or 'rows' not in data:
                self.send_error_response(400, "No data provided")
                return

            # Convert data back to DataFrame
            df = pd.DataFrame(data['rows'])

            # Generate file based on format
            if file_format == 'excel':
                file_content, content_type, file_extension = self.generate_excel(df)
            elif file_format == 'csv':
                file_content, content_type, file_extension = self.generate_csv(df)
            else:
                self.send_error_response(400, "Unsupported file format")
                return

            # Generate filename
            base_filename = os.path.splitext(original_filename)[0]
            filename = f"{base_filename}_CutFrame.{file_extension}"

            # Send file response
            self.send_file_response(file_content, content_type, filename)

        except Exception as e:
            self.send_error_response(500, f"Internal server error: {str(e)}")

    def generate_excel(self, df):
        """Generate Excel file"""
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='CutFrame')
        output.seek(0)
        
        return output.getvalue(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xlsx'

    def generate_csv(self, df):
        """Generate CSV file with proper column order matching original format"""
        # Define the expected column order from original CSV
        expected_columns = [
            'Batch No', 'Order No', 'Order Item', 'Material Name', 'Cutting ID', 'Pieces ID', 
            'Length', 'Angles', 'Qty', 'Bin No', 'Cart No', 'Position', 'Label Print', 
            'Barcode No', 'PO No', 'Style', 'Frame', 'Product Size', 'Color', 'Grid', 
            'Glass', 'Argon', 'Painting', 'Product Date', 'Balance', 'Shift', 'Ship date', 
            'Note', 'Customer'
        ]
        
        # Create a new DataFrame with the expected columns
        csv_df = pd.DataFrame()
        
        # Map existing columns to expected columns and preserve original data
        for col in expected_columns:
            if col in df.columns:
                csv_df[col] = df[col]
            else:
                # Fill missing columns with empty values or appropriate defaults based on original patterns
                if col == 'Batch No':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Order Item':
                    csv_df[col] = 1  # Default order item
                elif col == 'Angles':
                    csv_df[col] = 'V'  # Default angle value from original
                elif col == 'Cart No':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Position':
                    csv_df[col] = 'TOP+BOT'  # Most common position in original
                elif col == 'Label Print':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Barcode No':
                    csv_df[col] = ''  # Empty in original
                elif col == 'PO No':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Style':
                    csv_df[col] = 'XO'  # Most common style in original
                elif col == 'Frame':
                    csv_df[col] = 'Retrofit'  # Most common frame type in original
                elif col == 'Product Size':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Grid':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Glass':
                    csv_df[col] = 'Lowe3 Tmprd'  # Most common glass type in original
                elif col == 'Argon':
                    csv_df[col] = 'Argon'  # Common value in original
                elif col == 'Painting':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Product Date':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Balance':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Shift':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Ship date':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Note':
                    csv_df[col] = ''  # Empty in original
                elif col == 'Customer':
                    csv_df[col] = ''  # Will be empty without original data
                else:
                    csv_df[col] = ''  # Empty string for other missing columns
        
        # Ensure the DataFrame has the columns in the correct order
        csv_df = csv_df[expected_columns]
        
        # Generate CSV data
        csv_data = csv_df.to_csv(index=False)
        
        return csv_data.encode('utf-8'), 'text/csv', 'csv'

    def send_file_response(self, file_content, content_type, filename):
        """Send file response with proper headers"""
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Disposition', f'attachment; filename={filename}')
        self.end_headers()
        
        if isinstance(file_content, str):
            file_content = file_content.encode('utf-8')
        
        self.wfile.write(file_content)

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

    def generate_excel(self, df):
        """Generate Excel file from DataFrame"""
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='CutFrame')
        
        output.seek(0)
        content = output.getvalue()
        
        return content, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xlsx'

    def generate_csv(self, df):
        """Generate CSV file from DataFrame"""
        output = BytesIO()
        
        # Convert DataFrame to CSV
        csv_content = df.to_csv(index=False, encoding='utf-8')
        content = csv_content.encode('utf-8')
        
        return content, 'text/csv', 'csv'

    def send_file_response(self, content, content_type, filename):
        """Send file as response"""
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Content-Length', str(len(content)))
        self.end_headers()
        
        self.wfile.write(content)

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