from http.server import HTTPServer, SimpleHTTPRequestHandler
import subprocess
import os
from urllib.parse import urlparse, parse_qs

class LauncherHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed_path = urlparse(self.path)
        
        # Serve the main HTML file
        if self.path == '/':
            self.path = '/launcher.html'
            return SimpleHTTPRequestHandler.do_GET(self)
            
        # Handle script launching
        elif parsed_path.path == '/launch':
            query = parse_qs(parsed_path.query)
            script = query.get('script', [None])[0]
            
            if script and script.endswith('.py'):
                try:
                    # Launch the Python script
                    process = subprocess.Popen(['python', script], 
                                            stdout=subprocess.PIPE,
                                            stderr=subprocess.PIPE)
                    
                    self.send_response(200)
                    self.send_header('Content-type', 'text/plain')
                    self.end_headers()
                    self.wfile.write(f"Started {script}\n".encode())
                except Exception as e:
                    self.send_response(500)
                    self.send_header('Content-type', 'text/plain')
                    self.end_headers()
                    self.wfile.write(f"Error launching {script}: {str(e)}\n".encode())
            else:
                self.send_response(400)
                self.send_header('Content-type', 'text/plain')
                self.end_headers()
                self.wfile.write(b"Invalid script name\n")
                
        # Serve other files normally
        else:
            return SimpleHTTPRequestHandler.do_GET(self)

def run_server(port=8000):
    server_address = ('', port)
    httpd = HTTPServer(server_address, LauncherHandler)
    print(f"Server running on http://localhost:{port}")
    httpd.serve_forever()

if __name__ == '__main__':
    run_server()
