# data-visualization-project

Instructions: Put File in , then click the convert button

Install Instructions:

Step 1: Download and install Python https://www.python.org/downloads/

Step 2: Inside of the project create a python virtual environment using the following command: python -m venv environment-name (Replacing environment-name with whatever name you want)

Step 3: Run the virtual environment on windows in git bash using the following command: source environment-name/scripts/activate (Or on mac and linux source environment-name/bin/activate)

Step 4: Install all of the necessary packages using the requirements.txt with the following command: pip install -r requirements.txt

Step 5: Setup and run the flask server with the following commands:

export FLASK_ENV=development

export FLASK_APP=server.py

flask run
