import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pyvisa
import time
import csv

class IVAppCC:
    def __init__(self, root):
        self.root = root
        self.root.title("I-V Curve Measurement (CC Mode)")

        # Current input fields
        tk.Label(root, text="Start Current (A):").grid(row=0, column=0)
        self.start_current_entry = tk.Entry(root)
        self.start_current_entry.grid(row=0, column=1)

        tk.Label(root, text="End Current (A):").grid(row=1, column=0)
        self.end_current_entry = tk.Entry(root)
        self.end_current_entry.grid(row=1, column=1)

        tk.Label(root, text="Step Current (A):").grid(row=2, column=0)
        self.step_current_entry = tk.Entry(root)
        self.step_current_entry.grid(row=2, column=1)

        # Start button
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep)
        self.start_button.grid(row=3, column=0, columnspan=2)

        # Matplotlib figure for I-V plot
        self.figure = plt.Figure(figsize=(5, 4), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=4, column=0, columnspan=2)

    def start_sweep(self):
        try:
            # Parse user input
            i_start = float(self.start_current_entry.get())
            i_end = float(self.end_current_entry.get())
            i_step = float(self.step_current_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers for current values.")
            return

        # Connect to the instrument
        rm = pyvisa.ResourceManager()
        try:
            load = rm.open_resource("USB0::0x2EC7::0x8800::802199042787070066::INSTR")  
            load.write("MODE CC")
            load.write("INPUT ON")
        except Exception as e:
            messagebox.showerror("Connection Error", f"Could not connect to instrument: {e}")
            return

        # Perform current sweep
        currents = []
        voltages = []

        current = i_start
        while current <= i_end:
            load.write(f"CURR {current:.3f}")
            time.sleep(0.2)
            voltage = float(load.query("MEAS:VOLT?"))
            actual_current = float(load.query("MEAS:CURR?"))
            currents.append(actual_current)
            voltages.append(voltage)
            current += i_step

        load.write("INPUT OFF")

        # Save to CSV
        with open("iv_curve_gui_cc.csv", mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Current (A)", "Voltage (V)"])
            for i, v in zip(currents, voltages):
                writer.writerow([i, v])

        # Plot the I-V curve
        self.ax.clear()
        self.ax.plot(voltages, currents, marker='o')
        self.ax.set_title("I-V Curve (CC Mode)")
        self.ax.set_xlabel("Voltage (V)")
        self.ax.set_ylabel("Current (A)")
        self.ax.grid(True)
        self.canvas.draw()

        messagebox.showinfo("Success", "Sweep complete. Data saved to iv_curve_gui_cc.csv.")

if __name__ == "__main__":
    root = tk.Tk()
    app = IVAppCC(root)
    root.mainloop()
