# Setup Instructions
## IMPORTANT: Make sure that Python is installed in your machine. If not, install and set it up first.
## Download Pre-requisites
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

## Set Up and Run the Project

4. **Open a terminal and navigate to the project directory:**
   ```bash
   cd path/to/generate-reports
5. **Start Python Virtual Environment:**
   ```bash
   python3 -m venv venv
   ```
6. **Activate Virtual Environment::**
   ```bash 
   . venv/bin/activate
   ```
7. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
8. **Run Script:**
   ```bash
   python3 generate-reports.py reports.csv
   ```
