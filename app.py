from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
import os
import time
import threading
from selenium_bot.bot import CertificateBot
from utils.excel_reader import read_excel_row_by_row
# from werkzeug.middleware.dispatcher import DispatcherMiddleware
# from werkzeug.serving import run_simple

import logging

UPLOAD_FOLDER = 'uploads'
EXCEL_FOLDER = os.path.join(UPLOAD_FOLDER, 'excels')
DOWNLOAD_FOLDER = os.path.join(UPLOAD_FOLDER, 'downloads')

app = Flask(__name__)
app.secret_key = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Setup Logging
logging.basicConfig(filename='logs/bot.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Create upload folders if not exist
for folder in [UPLOAD_FOLDER, EXCEL_FOLDER, DOWNLOAD_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

bot_instance = None

def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)

def cleanup_old_logs(days=7):
    now = time.time()
    for filename in os.listdir("logs"):
        file_path = os.path.join("logs", filename)
        if os.path.isfile(file_path) and now - os.path.getmtime(file_path) > days * 86400:
            os.remove(file_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    global bot_instance

    if request.method == 'POST':
        clear_folder(EXCEL_FOLDER)

        excel_file = request.files.get('excel_file')
        username = request.form.get('username')
        password = request.form.get('password')

        if not excel_file or not username or not password:
            flash('Excel file, Username and Password are mandatory.', 'error')
            return redirect(request.url)

        # Save Excel
        excel_path = os.path.join(EXCEL_FOLDER, secure_filename(excel_file.filename))
        excel_file.save(excel_path)

        # Read Excel Data
        excel_data = read_excel_row_by_row(excel_path)

        # Create bot instance
        bot_instance = CertificateBot(
            username=username,
            password=password,
            excel_data=excel_data,
            download_folder=DOWNLOAD_FOLDER
        )

        # Start bot asynchronously to process all rows
        threading.Thread(target=start_bot).start()
        flash('Bot Started Successfully! Please monitor your downloads and logs.', 'success')
        return redirect(url_for('index'))

    return render_template('index.html')

def start_bot():
    try:
        result = bot_instance.process_all_certificates()
        if result.get("success"):
            logging.info("✅ All certificates processed successfully.")
        else:
            logging.error(f"❌ Error during processing: {result.get('message')}")
    except Exception as e:
        logging.error(f"❌ Unexpected error while processing certificates: {e}")

@app.route('/success')
def success():
    return "Automation Started Successfully! Check your downloads and logs."


# # Wrap the app to mount it under /e-payment
# application = DispatcherMiddleware(Flask('dummy'), {
#     '/e-payment': app
# })

if __name__ == '__main__':
    # Change host to '0.0.0.0' to listen on all interfaces, or use your specific IP
    app.run(host='0.0.0.0', port=5000, debug=True)
    # run_simple('0.0.0.0', 5050, application, use_debugger=True)
