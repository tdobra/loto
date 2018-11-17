#!/usr/bin/env python3

import http.server
import socketserver
import webbrowser
import threading

PORT = 8000

# Start server
Handler = http.server.SimpleHTTPRequestHandler
httpd = socketserver.TCPServer(("", PORT), Handler)
server_thread = threading.Thread(target=httpd.serve_forever)
# Exit the server thread when the main thread terminates
server_thread.daemon = True
server_thread.start()

# Display status
urlstr = "http://localhost:" + str(PORT)
print("Open " + urlstr + " in your web browser.")
print("Press Enter to exit.")
webbrowser.open(urlstr)

# Wait for user
input()

# Exit
httpd.shutdown()
httpd.server_close()
print("Program closed. You may need to wait a minute before reopening.")
