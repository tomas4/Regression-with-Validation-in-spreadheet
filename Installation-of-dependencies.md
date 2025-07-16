# **Installing scientific Python libraries pandas, scikit-learn, and numpy into LibreOffice's bundled Python environment**

The key is to always use the *specific Python executable bundled with LibreOffice* for all commands, not your system's general Python.  
---

### **General Steps (Applicable to All OS)**

1. **Find LibreOffice's Python Executable:** This is the most crucial step. You need to locate the python.exe (Windows), python3 (Linux), or python.app/Contents/MacOS/python (macOS) that comes with your LibreOffice installation.  
   * **Typical Locations:**  
     * **Windows:** C:\\Program Files\\LibreOffice\\program\\python.exe (or python.bat)  
     * **Linux:** /usr/lib/libreoffice/program/python or similar, depending on your distribution and installation method (e.g., Snap, Flatpak, system package). You might need to use which python3 or find /usr/lib \-name python3 after installing LibreOffice.  
     * **macOS:** /Applications/LibreOffice.app/Contents/Frameworks/Python.framework/Versions/\<Python\_Version\>/bin/python3 (or python)  
2. **Open a Command Prompt/Terminal:** You'll be typing commands directly.  
3. **Navigate to the LibreOffice Program Directory (Optional but Recommended):**  
   * Change your current directory to where LibreOffice's main executable and Python are located. This often simplifies commands, but using the full path to the Python executable works too.  
4. **Install/Update pip:** Even if pip seems to be there, ensuring it's up-to-date is vital.  
5. **Install Libraries with pip \--user:** Use the \--user flag to install packages into your user-specific site-packages directory, which does not require administrator privileges. LibreOffice's Python is usually configured to find packages installed this way.

---

### **Windows**
*Note: The instructions for Windows should **not** need administrative privileges to allow working also in resticted (typically corporate) Windows installations.*

**Assumptions:**

* LibreOffice is installed in the default location: C:\\Program Files\\LibreOffice.  
* You are using Command Prompt or PowerShell.

**Steps:**

1. **Open Command Prompt**  
   * Press Win \+ R, type cmd (for Command Prompt) or powershell (for PowerShell), and press Enter.  
2. **Download `get-pip.py` to temporary directory:**  
   * *Using Command Prompt:*

   DOS

   powershell \-Command "(Invoke-WebRequest \-Uri 'https://bootstrap.pypa.io/get-pip.py' \-UseBasicParsing).Content | Set-Content '%TEMP%\\get-pip.py'"

   * *Or manually:*

   If PowerShell is blocked or you prefer, you can manually download get-pip.py from [https://bootstrap.pypa.io/get-pip.py](https://bootstrap.pypa.io/get-pip.py) using your browser and save it to the directory indicated by typing   
     echo %TEMP%   
     in Command Prompt.

   

3. **Locate LibreOffice's Python Executable:**  
   * *It should be:*

   `"C:\Program Files\LibreOffice\program\python.exe"`

   *(Note: If python.exe is in a subfolder like python-core-3.x.x\\bin, adjust the path accordingly \- even in following commands.)*

4. **Run `get-pip.py` using LibreOffice's Python (from your Temporary directory):**  
   * Now, you'll execute the `get-pip.py` script that's in your `%TEMP%` folder, using LibreOffice's Python executable, and with the `--user` flag\> 

   

   DOS  
     "C:\\Program Files\\LibreOffice\\program\\python.exe" "%TEMP%\\get-pip.py" \--user

5. **Verify pip installation:**  
   DOS  
   "C:\\Program Files\\LibreOffice\\program\\python.exe" \-m pip \--version

   You should see output similar to pip 2x.x.x from C:\\Users\\\<YourUser\>\\AppData\\Roaming\\Python\\Python3x\\site-packages\\pip (python 3.x).  
6. **Install numpy, pandas, scikit-learn:**  
   DOS  
   "C:\\Program Files\\LibreOffice\\program\\python.exe" \-m pip install \--user numpy pandas scikit-learn

   This command will download and install these packages and their dependencies into your user's Python site-packages directory (e.g., C:\\Users\\\<YourUser\>\\AppData\\Roaming\\Python\\Python3x\\site-packages).

---

### **Linux**

**Assumptions:**

* LibreOffice's Python executable is typically found in /usr/lib/libreoffice/program/python3 or /usr/bin/libreoffice links to the correct script.  
* You are using a terminal (Bash, Zsh, etc.).

**Steps:**

1. **Open a Terminal:**  
   * Use your desktop environment's shortcut (e.g., Ctrl+Alt+T).  
2. **Determine LibreOffice's Python Path:**  
   * A common location for the Python executable bundled with LibreOffice is /usr/lib/libreoffice/program/python3 or /usr/lib/libreoffice/program/python.  
   * You can often verify its existence and version with:  
     Bash  
     /usr/lib/libreoffice/program/python3 \--version

     (If python3 isn't found, try python or explore ls \-l /usr/lib/libreoffice/program/.)  
   * Let's assume the path is /usr/lib/libreoffice/program/python3 for the following steps. Adjust if yours is different.  
3. **Install/Update pip:**  
   * Download get-pip.py:  
     Bash  
     wget https://bootstrap.pypa.io/get-pip.py \-O /tmp/get-pip.py

   * Run get-pip.py using LibreOffice's Python:  
     Bash  
     /usr/lib/libreoffice/program/python3 /tmp/get-pip.py \--user

4. **Verify pip installation:**  
   Bash  
   /usr/lib/libreoffice/program/python3 \-m pip \--version

   You should see output indicating pip is installed in your user's home directory (e.g., \~/.local/lib/python3.x/site-packages).  
5. **Install numpy, pandas, scikit-learn:**  
   Bash  
   /usr/lib/libreoffice/program/python3 \-m pip install \--user numpy pandas scikit-learn

   These packages will be installed in your user's local site-packages directory (e.g., \~/.local/lib/python3.x/site-packages).

---

### **macOS**

**Assumptions:**

* LibreOffice is installed in /Applications/LibreOffice.app.  
* You are using the Terminal app.

**Steps:**

1. **Open Terminal:**  
   * Go to Applications \> Utilities \> Terminal.  
2. **Determine LibreOffice's Python Path:**  
   * The Python executable is usually buried within the .app bundle. The path will look something like this (the version number will vary):  
     /Applications/LibreOffice.app/Contents/Frameworks/Python.framework/Versions/3.x/bin/python3  
   * To find your specific version:  
     Bash  
     ls \-d /Applications/LibreOffice.app/Contents/Frameworks/Python.framework/Versions/\*

     This will show the available Python versions. Pick the latest 3.x version.  
   * Let's assume the path is /Applications/LibreOffice.app/Contents/Frameworks/Python.framework/Versions/3.x/bin/python3 for the following steps. Replace 3.x with your actual version.  
3. **Install/Update pip:**  
   * Download get-pip.py:  
     Bash  
     curl https://bootstrap.pypa.io/get-pip.py \-o /tmp/get-pip.py

   * Run get-pip.py using LibreOffice's Python:  
     Bash  
     /Applications/LibreOffice.app/Contents/Frameworks/Python.framework/Versions/3.x/bin/python3 /tmp/get-pip.py \--user

4. **Verify pip installation:**  
   Bash  
   /Applications/LibreOffice.app/Contents/Frameworks/Python.framework/Versions/3.x/bin/python3 \-m pip \--version

   You should see output indicating pip is installed in your user's home directory (e.g., \~/Library/Python/3.x/lib/python/site-packages).  
5. **Install numpy, pandas, scikit-learn:**  
   Bash  
   /Applications/LibreOffice.app/Contents/Frameworks/Python.framework/Versions/3.x/bin/python3 \-m pip install \--user numpy pandas scikit-learn

   These packages will be installed in your user's local site-packages directory.

---

**Important Notes for All OS:**

* **Internet Connection:** These steps require an active internet connection to download get-pip.py and the library packages.  
* **Time:** Installing numpy, pandas, and especially scikit-learn (due to its SciPy dependency) can take a considerable amount of time as they often involve compiling C/Fortran code. Be patient.  
* **Errors:** If you encounter Permission Denied errors, double-check that you are using \--user and that you are not trying to install into system-wide LibreOffice directories directly.  
* **Python Version:** Always ensure the get-pip.py script you download is compatible with LibreOffice's *specific* Python version (e.g., Python 3.8, Python 3.9). The https://bootstrap.pypa.io/get-pip.py URL usually provides the latest compatible version, but for very old LibreOffice Python, you might need to find an archived version of get-pip.py (e.g., from https://bootstrap.pypa.io/pip/\<python-version\>/get-pip.py).  
* **Using the Libraries in LibreOffice:** After installation, you can then import and use numpy, pandas, and sklearn in your LibreOffice Python macros just like any other module.

[This video demonstrates how to install Python on a PC without local admin rights](https://www.youtube.com/watch?v=D15Gqsx8ffg), which is relevant to finding and using LibreOffice's bundled Python for package installation.  
Install Python on a locked down PC without local admin.  
