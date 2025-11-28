from flask import Flask, render_template, request, redirect, url_for, flash
import os
import time
import threading
from selenium_bot.bot import CertificateBot
import pandas as pd
from datetime import datetime
import logging

UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = os.path.join(UPLOAD_FOLDER, 'downloads')

app = Flask(__name__)
app.secret_key = 'your_secret_key'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Setup Logging
logging.basicConfig(filename='logs/bot.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Create upload folders if not exist
for folder in [UPLOAD_FOLDER, DOWNLOAD_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

bot_instance = None

def cleanup_old_logs(days=7):
    now = time.time()
    if os.path.exists("logs"):
        for filename in os.listdir("logs"):
            file_path = os.path.join("logs", filename)
            if os.path.isfile(file_path) and now - os.path.getmtime(file_path) > days * 86400:
                os.remove(file_path)

@app.route('/', methods=['GET', 'POST'])
def index():
    global bot_instance

    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        process_type = request.form.get('process_type')  # 'epcg', 'adv', or 'both'
        
        # Get EPCG data
        epcg_sb_date = request.form.get('epcg_sb_date')
        epcg_auth_no = request.form.get('epcg_auth_no')
        
        # Get ADV data
        adv_sb_date = request.form.get('adv_sb_date')
        adv_auth_no = request.form.get('adv_auth_no')

        if not username or not password:
            flash('Username and Password are mandatory.', 'error')
            return redirect(request.url)
        
        if not process_type:
            flash('Please select process type (EPCG, ADV, or Both).', 'error')
            return redirect(request.url)

        # Create Excel data structure
        excel_data = []
        row_data = {}
        
        if process_type in ['epcg', 'both']:
            if epcg_sb_date and epcg_auth_no:
                row_data["EPCG Shipping Bill Date"] = datetime.strptime(epcg_sb_date, '%Y-%m-%d')
                row_data["EPCG Authorisation Number"] = epcg_auth_no
            else:
                flash('EPCG Shipping Bill Date and Authorisation Number are required for EPCG process.', 'error')
                return redirect(request.url)
        
        if process_type in ['adv', 'both']:
            if adv_sb_date and adv_auth_no:
                row_data["ADV Shipping Bill Date"] = datetime.strptime(adv_sb_date, '%Y-%m-%d')
                row_data["ADV Authorisation Number"] = adv_auth_no
            else:
                flash('ADV Shipping Bill Date and Authorisation Number are required for ADV process.', 'error')
                return redirect(request.url)
        
        if row_data:
            excel_data.append(row_data)

        if not excel_data:
            flash('No valid data provided for processing.', 'error')
            return redirect(request.url)

        # Create bot instance
        bot_instance = CertificateBot(
            username=username,
            password=password,
            excel_data=excel_data,
            download_folder=DOWNLOAD_FOLDER,
            process_type=process_type
        )

        # Start bot asynchronously
        threading.Thread(target=start_bot).start()
        # flash('Bot Started Successfully! Please monitor your downloads and logs.', 'success')
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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)