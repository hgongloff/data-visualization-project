#!/bin/bash
source visualizer-env/Scripts/activate
export FLASK_ENV=development
export FLASK_APP=server.py
flask run