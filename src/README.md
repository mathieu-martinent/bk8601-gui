# BK8601-GUI

A Python-based GUI to control the **BK Precision 8601 DC Electronic Load**, allowing automated I-V curve measurements in **CC (Constant Current)** and **CV (Constant Voltage)** modes. The app is designed to work with real hardware or in a **simulated mode** for testing purposes.

---

## ✨ Features

* 🖥️ Clean and responsive GUI built with **Tkinter**
* 🔌 Compatible with **BK8601** via **PyVISA**, or use a simulated device if not connected
* 📊 Real-time plotting of **I-V** and **P-V** curves with **Matplotlib**
* 🧪 Supports both **2-Wire** and **4-Wire** sensing
* ⚙️ Selectable operation modes: **Constant Current (CC)** or **Constant Voltage (CV)**
* 🔐 Built-in protection: set current/voltage limits to avoid damaging the DUT (solar cell)
* 💾 Optional saving of data as **CSV** and plots as **PNG**
* 💡 Highlights Maximum Power Point (Pmp, Vmp, Imp)
* 💼 Saves last used settings to reload on next launch

---

## 📁 Project Structure

```bash
bk8601-gui/
├── src/
│   ├── main.py                # Main GUI entry point
│   ├── last_settings.json     # Stores previous session settings
│   ├── manual_tests/          # Manual test files and planning
│   │   └── CC-CV_modes_plan.ods
├── requirements.txt          # List of dependencies
└── output/                   # Automatically generated measurement folders
```

---

## 🚀 Installation & Usage

### 1. Clone the repository

```bash
git clone https://github.com/mathieu-martinent/bk8601-gui.git
cd bk8601-gui/src
```

### 2. Create a virtual environment (recommended)

```bash
python -m venv .venv
# Then activate it:
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate
```

### 3. Install dependencies

```bash
pip install -r ../requirements.txt
```

### 4. Run the application

```bash
python main.py
```

> ⚠️ **No BK8601?** The app includes a built-in simulated instrument so you can test the GUI even without the hardware connected.

---

## 🧪 Manual Tests

You’ll find manual test files and mode planning in:

```
src/manual_tests/CC-CV_modes_plan.ods
```

This spreadsheet helps document different test configurations (2W/4W, CC/CV).

---

## 🔮 Future Improvements (TODO)

* Add automated test scripts for regression testing
* Improve error handling for VISA exceptions
* Export data in JSON format for further analysis
* Add GUI language selector (EN/FR)

---

## 📝 License

MIT License — feel free to use, modify, and share.

---

## 👤 Author

**Mathieu Martinent**

> Built as part of a technical internship focused on embedded instrumentation and solar cell testing automation.
