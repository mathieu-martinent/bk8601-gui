# BK8601-GUI

Graphical user interface (GUI) in Python for automated I-V curve measurement using the BK Precision 8601 DC electronic load.

## 📌 Description

This project provides an interactive tool to control the BK Precision 8601 for current/voltage sweep measurements on electronic devices (e.g., solar cells).  
The GUI allows configuration of the test mode, sweep parameters, safety limits, and live plotting of results.

The application supports:
- **CC mode** (Constant Current)
- **CV mode** (Constant Voltage)
- **2-wire and 4-wire sensing**
- **Live plot of I-V and Power curves**
- **CSV and PNG export of results**

## 🛠️ Technologies

- Python 3
- `tkinter` for GUI
- `matplotlib` for plotting
- `pyvisa` for instrument communication

## 📂 Project structure

```
bk8601-gui/
├── src/                # Core GUI and measurement logic
│   └── main.py
├── output/             # Measurement results (.csv, .png) – ignored by Git
├── tests/              # Test scripts and early versions
├── .gitignore
├── README.md
```

## ▶️ How to run

1. Clone the repository:
   ```bash
   git clone https://github.com/mathieu-martinent/bk8601-gui.git
   cd bk8601-gui
   ```

2. Install required packages:
   ```bash
   pip install matplotlib pyvisa
   ```

3. Run the GUI:
   ```bash
   python src/main.py
   ```

## 📊 Example output

- CSV: I-V-P sweep data
- PNG: Exported plot with I-V and Power curves
- Summary: Maximum power point (Pmp), voltage (Vmp), current (Imp)

## 🧪 Simulated instrument support

If no real device is connected, the GUI offers a **simulated BK8601** to allow offline testing and debugging.

## 🧑‍💻 Author

**Mathieu Martinent**  
Electrical & Control Engineering Student – ENSEA  
[LinkedIn] www.linkedin.com/in/mathieu-martinent

---

Feel free to fork or contribute. Feedback and improvements welcome!
