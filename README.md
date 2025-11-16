# QPrimeLauncher

QPrimeLauncher is a Windows-based launcher and updater utility designed to ensure the correct and latest version of the **QPRIME Client** is installed and running. It is a modernised, PowerShell-driven replacement for the legacy **QPrimeLaunch.vbs** script used within QPS environments.

The goal of QPrimeLauncher is to provide:
- Reliable version checking  
- Clean update logic  
- Clear user prompts and error messages  
- Centralised logging  
- Support for SCCM/SMS deployments  
- Maintainable, testable, modern PowerShell code  

---

## 🚀 Features

- **INI-driven configuration**  
  Reads remote INI files to determine the latest valid client version.

- **Full update workflow**  
  Detects version mismatches and runs the updater with correct flags.

- **Graceful error handling**  
  Includes robust try/catch blocks, logging, and fallback logic.

- **Logging output**  
  Generates timestamped logs for auditing and troubleshooting.

- **User prompts**  
  Displays friendly messages when critical action is required.

- **SCCM/SMS compatible**  
  Designed to integrate with enterprise deployment pipelines.

- **Modern PowerShell architecture**  
  Replaces VBScript with structured functions, modularity, and maintainability.

---

## 📁 Project Structure

QPrimeLauncher/
├── QPrimeLauncher.ps1 # Main entry script
├── modules/
│ ├── GetLatestVersion.psm1 # Remote INI version extraction
│ ├── Logging.psm1 # Log writer + formatting
│ ├── Utils.psm1 # Helpers, path checks, version parsing
│ └── Updater.psm1 # Runs QPRIME updater safely
├── config/
│ └── settings.ini # Sample local INI for testing
├── logs/ # Generated automatically
└── README.md

yaml
Copy code

---

## 🔧 How It Works

1. **Load Configuration**  
   The script reads the remote INI file from a UNC share to determine the correct version.

2. **Compare Versions**  
   Local client version is compared with the required version defined in the INI.

3. **Decide Path**  
   - ✔️ If up-to-date → client launches  
   - ❗ If outdated → updater is executed

4. **Run Updater**  
   Executes the QPRIME updater via PowerShell with correct flags and error checking.

5. **Write Logs**  
   Every action is logged for diagnostics.

6. **Launch Application**  
   The correct QPRIME client starts automatically.

---

## 🛠 Requirements

- Windows 10 or later  
- PowerShell 5.1  
- Access to QPRIME network shares  
- Permissions to run updater executables  
- SCCM/SMS optional

---

## 📦 Installation & Usage

### **Run the launcher:**

```powershell
.\QPrimeLauncher.ps1
