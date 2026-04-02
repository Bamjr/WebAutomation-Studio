# WebAutomatic_GUI(BETA)
This tool is designed to streamline and simplify the process for users who wish to automate tasks using a bot. (WIP)

## 🚀 How to Use

1. Make sure you have Python installed. (recommend 3.11.9 version)
2. Install dependencies (e.g., `selenium`, `ttkbootstrap`) via:

   ```bash
   pip install selenium
   python -m pip install ttkbootstrap

   ```

3. Run the script [Auto.py](Auto.py):

   ```bash
   python Auto.py
   ```

4. Use the GUI to:
   - Select your Chrome profile folder
   - Choose your `chrome.exe` and `chromedriver.exe`
   - (Optional) Enter your Chrome version
   - Click **"Start"** to launch Chrome with custom configuration

---

## ✅ Current Status

- [x] GUI implemented with `tkinter` and `ttkbootstrap`
- [x] Able to launch **custom Chrome instance** using user-defined:
  - Profile folder
  - Chrome binary
  - ChromeDriver path
- [x] Optional Chrome version input supported
- [x] Configuration fully integrated with Selenium
- [ ] on going - Action function (find, click etc.)

---
