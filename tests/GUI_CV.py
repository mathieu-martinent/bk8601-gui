import pyvisa

# Create a ResourceManager object to manage connected instruments
rm = pyvisa.ResourceManager()

# List all connected instruments
instruments = rm.list_resources()

# Display the addresses of all detected instruments
print("Connected instruments:")
for instrument in instruments:
    print(instrument)

import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pyvisa
import time

class IVApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IV Curve Measurement")

        # Voltage input fields
        tk.Label(root, text="Start Voltage (V):").grid(row=0, column=0)
        self.start_voltage_entry = tk.Entry(root)
        self.start_voltage_entry.grid(row=0, column=1)

        tk.Label(root, text="End Voltage (V):").grid(row=1, column=0)
        self.end_voltage_entry = tk.Entry(root)
        self.end_voltage_entry.grid(row=1, column=1)

        tk.Label(root, text="Step Voltage (V):").grid(row=2, column=0)
        self.step_voltage_entry = tk.Entry(root)
        self.step_voltage_entry.grid(row=2, column=1)

        # Start button
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep)
        self.start_button.grid(row=3, column=0, columnspan=2)

        # Placeholder for matplotlib figure
        self.figure = plt.Figure(figsize=(5, 4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=4, column=0, columnspan=2)

    def start_sweep(self):
        try:
            # Retrieve and convert user input for start, end, and step voltages
            v_start = float(self.start_voltage_entry.get())
            v_end = float(self.end_voltage_entry.get())
            v_step = float(self.step_voltage_entry.get())
        except ValueError:
             # Show an error message if the input is invalid
            messagebox.showerror("Input Error", "Please enter valid numbers for voltage values.")
            return

        # Connect to instrument
        rm = pyvisa.ResourceManager()
        try:
            # Connect to the instrument using its VISA address
            load = rm.open_resource("USB0::0x2EC7::0x8800::802199042787070066::INSTR")  
            # Set the instrument to constant voltage (CV) mode and enable input
            load.write("MODE CV")
            load.write("INPUT ON")
        except Exception as e:
            # Show an error message if the connection fails
            messagebox.showerror("Connection Error", f"Could not connect to instrument: {e}")
            return

        # Initialize lists to store voltage and current measurements
        voltages = []
        currents = []
        v = v_start
        while v <= v_end:
            load.write(f"VOLT {v:.3f}")
            time.sleep(0.2)
             # Query the current measurement from the instrument
            i = float(load.query("MEAS:CURR?"))
            # Append the voltage and current to their respective lists
            voltages.append(v)
            currents.append(i)
            v += v_step

        load.write("INPUT OFF")

        # Plot results
        self.ax.clear()
        self.ax.plot(voltages, currents, marker='o')
        self.ax.set_title("I-V Curve")
        self.ax.set_xlabel("Voltage (V)")
        self.ax.set_ylabel("Current (A)")
        self.canvas.draw()

if __name__ == "__main__":
    root = tk.Tk()
    app = IVApp(root)
    root.mainloop()
