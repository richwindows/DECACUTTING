#!/usr/bin/env python3
"""
Local development server for testing the DECA Cutting application
This server can handle both static files and API endpoints
"""

import os
import sys
import socket
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
import urllib.parse
import json
import traceback

# Add the current directory to Python path so we can import our API modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

class LocalDevHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, directory=None, **kwargs):
        if directory is None:
            directory = os.getcwd()
        super().__init__(*args, directory=directory, **kwargs)
    
    def do_POST(self):
        """Handle POST requests for API endpoints"""
        if self.path.startswith('/api/'):
            self.handle_api_request()
        else:
            self.send_error(404, "Not Found")
    
    def handle_api_request(self):
        """Handle API requests by forwarding to the appropriate handler"""
        try:
            if self.path == '/api/process':
                # Import the process handler
                from api.process import handler as ProcessHandler
                
                # Create a new handler instance
                api_handler = ProcessHandler(self.request, self.client_address, self.server)
                
                # Copy all necessary attributes
                for attr in ['rfile', 'wfile', 'headers', 'command', 'path', 'request_version', 'requestline']:
                    if hasattr(self, attr):
                        setattr(api_handler, attr, getattr(self, attr))
                
                # Process the request
                api_handler.do_POST()
                
            elif self.path == '/api/download':
                # Import the download handler
                from api.download import handler as DownloadHandler
                
                # Create a new handler instance
                api_handler = DownloadHandler(self.request, self.client_address, self.server)
                
                # Copy all necessary attributes
                for attr in ['rfile', 'wfile', 'headers', 'command', 'path', 'request_version', 'requestline']:
                    if hasattr(self, attr):
                        setattr(api_handler, attr, getattr(self, attr))
                
                # Process the request
                api_handler.do_POST()
                
            else:
                self.send_error(404, "API endpoint not found")
                
        except ConnectionAbortedError:
            print(f"Connection aborted by client during {self.path}")
        except Exception as e:
            print(f"Error handling {self.path}: {e}")
            traceback.print_exc()
            try:
                self.send_error(500, f"Internal server error: {e}")
            except ConnectionAbortedError:
                print("Connection aborted while sending error response")

    def do_GET(self):
        """Handle GET requests for static files"""
        try:
            super().do_GET()
        except ConnectionAbortedError:
            # Ignore connection aborted errors (common in development)
            pass
        except Exception as e:
            print(f"Error serving static file: {e}")

    def do_OPTIONS(self):
        """Handle preflight requests"""
        self.send_response(200)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def log_message(self, format, *args):
        """Override to reduce verbose logging"""
        # Only log errors and important requests
        if "404" in str(args) or "500" in str(args) or "POST" in str(args):
            super().log_message(format, *args)

class RobustHTTPServer(HTTPServer):
    """A more robust HTTP server that handles connection errors gracefully"""
    
    def handle_error(self, request, client_address):
        """Handle errors more gracefully"""
        try:
            super().handle_error(request, client_address)
        except Exception:
            # Ignore errors in error handling
            pass

if __name__ == "__main__":
    PORT = 8000
    
    print(f"Starting local development server on port {PORT}...")
    print(f"Server will serve files from: {os.getcwd()}")
    print(f"Available at: http://localhost:{PORT}")
    print("Press Ctrl+C to stop the server")
    
    try:
        with RobustHTTPServer(("", PORT), LocalDevHandler) as httpd:
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
    except Exception as e:
        print(f"Server error: {e}")
        traceback.print_exc()