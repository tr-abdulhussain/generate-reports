# Setup Instructions
## IMPORTANT: Make sure that Python is installed in your machine. If not, install and set it up first.
## Pre-requisites
1. **Download the Repository**  
   Clone or download the repository to your local machine.

2. **Download Reports from Snipe-IT**  
   - Download the latest asset reports from Snipe-IT.  
   - Rename the downloaded file to `reports.csv`.  
   - Move the file into the `generate-reports` directory.

3. **Download OAuth Client Credentials**  
   - Go to the [OAuth Client Credentials page](https://console.cloud.google.com/apis/credentials?hl=en&invt=AbutXg&project=automated-dependency-gathering).  
   - Download the Generate Machine Reports OAuth Client JSON file.  
   - Rename it to `credentials.json`.  
   - Move the file into the `generate-reports` directory.

4. **Create API Key in Snipe-IT**
   - Go to Snipe-IT, and click Manage API Keys under your Name.
   - Click Create Token, input the Name, then Create.
   - In terminal, save the token as `SNIPEIT_API_KEY`.
     ```bash
     export SNIPEIT_API_KEY=PASTE TOKEN HERE
     ```
   - Save your token somewhere safe like 1Password.

## Set Up and Run the Project

5. **Open a terminal and navigate to the project directory:**
   ```bash
   cd path/to/generate-reports
6. **Start Python Virtual Environment:**
   ```bash
   python3 -m venv venv
   ```
7. **Activate Virtual Environment::**
   ```bash 
   . venv/bin/activate
   ```
8. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
9. **Run Script:**
   ```bash
   python3 functions/generate-reports.py reports.csv
   ```
