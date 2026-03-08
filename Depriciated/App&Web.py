import subprocess
import shutil
from urllib.parse import quote_plus
import os


# ========================================================
# 1. OPEN ANY APPLICATION (EXE, LNK, FOLDER, ETC.)
# ========================================================
def open_app(path):
    """
    Opens any application or file:
        - .exe → subprocess.Popen
        - .lnk → os.startfile
    Returns True if successful, False otherwise.
    """
    try:
        path = os.path.normpath(path)

        # Handle .lnk shortcut
        if path.lower().endswith(".lnk"):
            try:
                os.startfile(path)
                return True
            except Exception as e:
                print("Failed to open shortcut:", e)
                return False

        # Handle .exe
        if path.lower().endswith(".exe"):
            try:
                subprocess.Popen([path])
                return True
            except Exception as e:
                print("Failed to open executable:", e)
                return False

    except Exception as e:
        print("Unexpected error:", e)
        return False


# ========================================================
# 2. FIND AN APPLICATION AUTOMATICALLY (OPTIONAL)
# ========================================================
def find_app(app_name):
    """
    Uses shutil.which() to auto-detect an installed system application.
    Example:
        find_app("chrome") → C:/Program Files/.../chrome.exe OR None
    """
    return shutil.which(app_name)


# ========================================================
# 3. SEARCH ANYTHING IN GOOGLE USING CHROME
# ========================================================
def chrome_search(query):
    """
    Opens Google Chrome and performs a Google search.
    Automatically detects chrome.exe when possible.
    """

    # Try auto-detecting Chrome
    chrome_path = find_app("chrome")

    # If not found, try common installation path
    if not chrome_path:
        chrome_candidates = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ]

        for candidate in chrome_candidates:
            if os.path.exists(candidate):
                chrome_path = candidate
                break

    # If Chrome still not found → fail gracefully
    if not chrome_path or not os.path.exists(chrome_path):
        print("Chrome not found on your system.")
        return False

    # Encode search text
    encoded = quote_plus(query)
    url = f"https://www.google.com/search?q={encoded}"

    # Launch Chrome with search URL
    try:
        subprocess.Popen([chrome_path, url])
        return True
    except Exception as e:
        print("Error launching Chrome search:", e)
        return False


# ========================================================
# 4. EXAMPLES (COMMENTED)
# ========================================================
"""
# Open Notepad
open_app(r"C:\\Windows\\System32\\notepad.exe")

# Open Microsoft Excel
open_app(r"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")

# Open a .lnk shortcut
open_app(r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Excel.lnk")

# Chrome Search
chrome_search("python list comprehension tutorial")


