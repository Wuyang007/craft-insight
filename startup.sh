#!/bin/bash
# Navigate to the site folder
cd $HOME/site/wwwroot

# Install dependencies
pip install --upgrade pip
pip install -r requirements.txt

# Start Flask app
gunicorn --bind=0.0.0.0 --timeout 600 app_v1:app