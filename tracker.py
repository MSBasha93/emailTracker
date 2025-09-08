# tracker.py
from flask import Flask, send_file, request
from datetime import datetime
import os
import io

app = Flask(__name__)

LOG_FILE = 'opens_log.csv'
PIXEL_BYTES = b'GIF89a\x01\x00\x01\x00\x80\x00\x00\xff\xff\xff\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;'

@app.route('/track/<tracking_id>')
def track_open(tracking_id):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    user_agent = request.headers.get('User-Agent', 'Unknown')
    
    # Write to the log file in CSV format
    with open(LOG_FILE, 'a') as f:
        f.write(f'"{timestamp}","{tracking_id}","{user_agent}"\n')
        
    print(f"✅ Tracked open for ID: {tracking_id}")

    return send_file(io.BytesIO(PIXEL_BYTES), mimetype='image/gif')

if __name__ == '__main__':
    # Make sure the log file exists with a header
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, 'w') as f:
            f.write("opened_time,tracking_id,user_agent\n")
            
    # Run the server on port 5000
    app.run(host='0.0.0.0', port=5000)