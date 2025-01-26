# Manage Outlook New

![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
<!-- PROJECT LOGO --> 
<div align="center"> <img src="images/icon.png" alt="ManageOutlook" width="80" height="80"><h3 align="center">Manage Outlook New</h3> <p align="center"> 
## Overview
The `manage_outlook_new.py` script is designed to manage the installation and configuration of Outlook New on Windows systems. It allows you to disable Outlook New, uninstall it, block future installations, and revert all changes made. Additionally, it can disable the built-in mail application in Windows and the Microsoft Store.

## Requirements
- Operating System: Windows 8 or higher.
- Administrator permissions.
- Python 3.x installed.

## Usage Instructions
The script can be run from the command line with different options to perform various actions. Below are the available options:

### Options
- `1`: Disable Outlook New.
- `2`: Uninstall Outlook New.
- `3`: Block future installations of Outlook New.
- `4`: Revert all changes made.
- `6`: Block reinstallation from Microsoft Store.
- `7`: Schedule a task to remove Outlook New periodically.
- `8`: Disable the built-in mail application in Windows.
- `9`: Update local group policies.
- `10`: Uninstall the built-in mail application in Windows.
- `11`: Disable the Microsoft Store completely.
- `12`: Uninstall the Microsoft Store application.
- `/?`: Show this help message.

### Examples
- To disable Outlook New:
  ```sh
  python manage_outlook_new.py 1
- To uninstall Outlook New:
  ```sh
  python manage_outlook_new.py 2
- To revert all changes made:
  ```sh
  python manage_outlook_new.py 4
- To show the help message:
  ```sh
  python manage_outlook_new.py /?
### Behavior When No Parameters Are Provided

When no parameters are added while running the script manage_outlook_new.py, the script will perform a series of predetermined actions. These actions are designed to disable and uninstall Outlook New, block future installations, disable the built-in mail application in Windows, and disable the Microsoft Store. Below is a step-by-step explanation of the behavior:
### Behavior Without Parameters
Behavior Without Parameters
Check for Administrator Permissions:

The script first checks if it is running with administrator permissions. If not, the script will stop and display a message indicating that administrator permissions are required.
Check Windows Version:

The script checks if the operating system is Windows XP or Windows 7. If so, the script will stop and display a message indicating that it cannot run on these versions of Windows.
Execute Predetermined Actions:

If no parameters are provided, the script will execute the following predetermined actions:
Disable Outlook New: Modifies a registry key to disable Outlook New.
Uninstall Outlook New: Uses PowerShell to uninstall Outlook New if it is installed.
Block Future Installations of Outlook New: Modifies a registry key to block future installations of Outlook New.
Disable the Built-in Mail Application in Windows: Uses PowerShell to disable the built-in mail application in Windows.
Block Reinstallation from Microsoft Store: Modifies registry keys to block the reinstallation of applications from the Microsoft Store.
Disable the Microsoft Store: Modifies a registry key to disable the Microsoft Store completely.
Uninstall the Built-in Mail Application in Windows: Uses PowerShell to uninstall the built-in mail application in Windows.
Uninstall the Microsoft Store Application: Uses PowerShell to uninstall the Microsoft Store application.
Schedule a Task to Remove Outlook New Periodically: Uses the Windows Task Scheduler to schedule a task that removes Outlook New periodically.
Update Local Group Policies: Uses gpupdate to update local group policies and apply the changes immediately.
