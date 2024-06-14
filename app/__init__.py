from flask import Flask
from pymongo import MongoClient
from gridfs import GridFS
import os

app = Flask(__name__, static_url_path='/static')

# Set the secret key for the app
app.secret_key = 'your_secret_key'

# Configure logging
import logging
logging.basicConfig(level=logging.DEBUG)

# MongoDB setup
client = MongoClient('mongodb+srv://beingchirag0051:Chirag5151@cluster0.e6xh9gf.mongodb.net/your_database?retryWrites=true&w=majority')
db = client['user_login_system']
fs = GridFS(db)

# Constants
COVER_LETTER_TEMPLATE = os.path.join(app.root_path, 'Cover_letterr.docx')
FINAL_FILE_DOCX_FILENAME = "Final_Cover_letter_with_table_{}.docx"
FINAL_FILE_PDF_FILENAME = "Final_Cover_letter_with_table_{}.pdf"
PDFS_DIRECTORY = os.path.join(app.root_path, 'app', 'pdfs')

# Ensure the PDFs directory exists
if not os.path.exists(PDFS_DIRECTORY):
    os.makedirs(PDFS_DIRECTORY)

# Import routes
from app import routes
