#!/usr/bin/env python3
"""
Simple Flask-based development server for testing the DECA Cutting application
"""

import os
import sys
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import tempfile
import traceback

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes

@app.route('/')
def index():
    """Serve the main HTML file"""
    return send_from_directory('.', 'index.html')

@app.route('/<path:filename>')
def static_files(filename):
    """Serve static files"""
    return send_from_directory('.', filename)

@app.route('/api/process', methods=['POST', 'OPTIONS'])
def api_process():
    """Handle file processing requests"""
    if request.method == 'OPTIONS':
        return '', 200
    
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'success': False, 'message': 'No file uploaded'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'success': False, 'message': 'No file selected'}), 400
        
        process_type = request.form.get('processType', 'Windows')
        
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1]) as tmp_file:
            file.save(tmp_file.name)
            tmp_file_path = tmp_file.name
        
        try:
            # Import and use the appropriate converter
            if process_type == 'Windows':
                import convertWindow
                df, _ = convertWindow.process_file(tmp_file_path)
            else:  # Door
                import convertDoor
                df, _ = convertDoor.process_file(tmp_file_path)
            
            # Process cutting data
            from app import process_cutting_data
            success, message, result_df = process_cutting_data(df)
            
            if not success:
                return jsonify({'success': False, 'message': message}), 400
            
            # Convert DataFrame to the expected format with JSON serialization fix
            import numpy as np
            import pandas as pd
            
            # Convert numpy types to native Python types for JSON serialization
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
                'filename': file.filename
            }
            
            return jsonify(response_data)
            
        finally:
            # Clean up temporary file
            if os.path.exists(tmp_file_path):
                os.unlink(tmp_file_path)
                
    except Exception as e:
        print(f"Error processing file: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Internal server error: {str(e)}'}), 500

@app.route('/api/download', methods=['POST', 'OPTIONS'])
def api_download():
    """Handle file download requests"""
    if request.method == 'OPTIONS':
        return '', 200
    
    try:
        data = request.get_json()
        
        if not data or 'data' not in data:
            return jsonify({'success': False, 'message': 'No data provided'}), 400
        
        # Import pandas for file generation
        import pandas as pd
        from io import BytesIO
        
        # Convert data back to DataFrame
        df = pd.DataFrame(data['data']['rows'])
        
        file_format = data.get('format', 'excel')
        original_filename = data.get('filename', 'processed_data')
        
        # Generate file based on format
        if file_format == 'excel':
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='CutFrame')
            output.seek(0)
            
            base_filename = os.path.splitext(original_filename)[0]
            filename = f"{base_filename}_CutFrame.xlsx"
            
            from flask import Response
            return Response(
                output.getvalue(),
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={'Content-Disposition': f'attachment; filename={filename}'}
            )
        elif file_format == 'csv':
            # Generate CSV file with proper column order matching original format
            
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
            
            base_filename = os.path.splitext(original_filename)[0]
            filename = f"{base_filename}_CutFrame.csv"
            
            from flask import Response
            return Response(
                csv_data,
                mimetype='text/csv',
                headers={'Content-Disposition': f'attachment; filename={filename}'}
            )
        else:
            return jsonify({'success': False, 'message': 'Unsupported file format'}), 400
            
    except Exception as e:
        print(f"Error generating download: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Internal server error: {str(e)}'}), 500

if __name__ == '__main__':
    PORT = 8000
    print(f"Starting Flask development server on port {PORT}...")
    print(f"Available at: http://localhost:{PORT}")
    print("Press Ctrl+C to stop the server")
    
    app.run(host='0.0.0.0', port=PORT, debug=True, use_reloader=False)