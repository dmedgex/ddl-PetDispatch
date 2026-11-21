# Installation Guide (Detailed)

This guide is designed for users who may not be familiar with Python or the command line. We recommend using **Conda** to manage your environment, which prevents conflicts with other software on your computer.

## Step 1: Install Miniconda
Conda is a tool that creates a "safe space" (environment) for this calculator to run in.

1.  **Download Miniconda**: 
    -   [Windows Installer](https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe)
    -   [macOS Installer](https://repo.anaconda.com/miniconda/Miniconda3-latest-MacOSX-x86_64.pkg)
    -   [Linux Installer](https://repo.anaconda.com/miniconda/Miniconda3-latest-Linux-x86_64.sh)
2.  **Follow the Installation Wizard**: Use the default settings. On Windows, make sure to search for **"Anaconda Prompt"** or **"Miniconda Prompt"** in your Start Menu after installation.

## Step 2: Download the Project
1.  Download the project as a ZIP file from GitHub.
2.  Extract the ZIP file to a folder you can easily find (e.g., your Desktop or Documents).

## Step 3: Set Up the Environment
Open your terminal (Anaconda/Miniconda Prompt on Windows, Terminal on macOS/Linux) and navigate to the project folder:

```bash
# Example for Windows (replace with your actual folder path)
cd C:\Users\YourName\Desktop\trickcal-team-optimizer
```

Once inside the folder, run this **one-liner** to create the environment and install everything automatically:

```bash
conda create -n petdispatch python=3.10 -y && conda activate petdispatch && pip install -r requirements.txt
```

## Step 4: Run the Calculator
Every time you want to use the calculator, open your prompt/terminal and run:

```bash
# 1. Activate the environment (only needs to be done once per session)
conda activate petdispatch

# 2. Run the Web Interface
streamlit run src/ui/web_gui.py
```

*This will open a new tab in your default web browser where you can use the calculator.*

---

## Alternative: Command Line Version
If you prefer the text-based version, run:
```bash
python main.py
```

## Troubleshooting
-   **"Command not found"**: Ensure you are using the **Anaconda/Miniconda Prompt** (on Windows) or have restarted your terminal after installing Miniconda.
-   **Wrong Folder**: Make sure you use the `cd` command to enter the folder where `main.py` and `requirements.txt` are located before running the setup command.
