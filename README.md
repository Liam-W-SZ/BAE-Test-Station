# BAE-Test-Station

## Overview
The **BAE Project** is a Python-based application designed to manage alarms, configurations, SharePoint integrations, and other utilities. It provides a modular structure with JSON-based configuration files and a virtual environment for dependency management. This project is ideal for automating tasks and managing data in a structured and efficient way.

---

## Features
- **Alarm Management**: Configure and manage alarms using `BAE_Alarms.json`.
- **SharePoint Integration**: Seamless integration with SharePoint, configured via `BAE_SharePoint_Config.json`.
- **Custom Tools**: Utility functions provided in `tv_tools.py` for additional operations.
- **JSON-Based Configurations**: Easily customizable settings for file paths, login credentials, and parameters.
- **Virtual Environment**: Pre-configured virtual environment for dependency isolation.

---

## Project Structure

### Root Files
- **`BAE_Alarms.json`**: Alarm configurations.
- **`BAE_Config_Parameters.json`**: Application parameters.
- **`BAE_File_Paths.json`**: File path configurations.
- **`BAE_Login.json`**: Login credentials or settings.
- **`BAE_SharePoint_Config.json`**: SharePoint integration settings.
- **`BAE_SW_Code.py`**: Main Python script for the application.
- **`tv_tools.py`**: Utility script for additional tools.
- **`start_bae.sh`**: Shell script to start the application.
- **`Help.txt`**: Documentation or help file.

### Virtual Environment (`VenvBAE/`)
- **`pyvenv.cfg`**: Virtual environment configuration.
- **`Scripts/`**: Activation and deactivation scripts.
- **`Lib/site-packages/`**: Installed Python packages.

---

## How It Works

1. **Configuration**:
   - The program uses JSON files to store configurations for alarms, file paths, SharePoint settings, and login credentials.
   - These files can be edited to customize the application's behavior.

2. **Main Script**:
   - The core logic resides in `BAE_SW_Code.py`, which reads the configuration files and executes the required tasks.
   - It interacts with SharePoint, manages alarms, and performs other operations based on the provided configurations.

3. **Utility Tools**:
   - Additional functionalities are provided in `tv_tools.py`, which can be used to extend the application's capabilities.

4. **Virtual Environment**:
   - The project includes a pre-configured virtual environment (`VenvBAE/`) to ensure all dependencies are isolated and compatible.

---

## Prerequisites
- Python 3.12 or higher
- Git (for cloning the repository)

---

## Installation

1. **Clone the Repository**:
   ```bash
   git clone <repository-url>
   cd BAE_Code
   ```

2. **Activate the Virtual Environment**:
   - On Windows:
     ```cmd
     VenvBAE\Scripts\activate.bat
     ```
   - On Linux/Mac:
     ```bash
     source VenvBAE/Scripts/activate
     ```

3. **Install Dependencies**:
   If additional dependencies are required, install them using:
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the Application**:
   Execute the main script:
   ```bash
   python BAE_SW_Code.py
   ```

---

## Configuration

### JSON Files
- **`BAE_Alarms.json`**: Define alarms and their triggers.
- **`BAE_Config_Parameters.json`**: Set application parameters.
- **`BAE_File_Paths.json`**: Specify file paths used by the application.
- **`BAE_Login.json`**: Configure login credentials.
- **`BAE_SharePoint_Config.json`**: Set SharePoint integration details.

### Example Configuration
Here’s an example of how a JSON configuration file might look:
```json
{
  "alarm_name": "Critical Alarm",
  "trigger": "Disk Space Low",
  "action": "Send Email Notification"
}
```

---

## Key Dependencies
The project uses the following Python libraries:
- **`numpy`**: For numerical computations.
- **`pandas`**: For data manipulation and analysis.
- **`matplotlib`**: For data visualization.
- **`customtkinter`**: For creating modern graphical user interfaces.
- **`pycurl`**: For making HTTP requests.
- **`cryptography`**: For secure data handling.
- **`openpyxl`**: For working with Excel files.

---

## Support
For any issues or questions, refer to the `Help.txt` file or open an issue on GitHub.

---

## Contact
For further inquiries, please contact the project maintainer.
