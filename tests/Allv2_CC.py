import tkinter as tk
from tkinter import messagebox, filedialog
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
        tk.Label(root, text="Start Current (A):").grid(row=0, column=0, sticky="e")
        self.start_current_entry = tk.Entry(root)
        self.start_current_entry.grid(row=0, column=1)

        tk.Label(root, text="End Current (A):").grid(row=1, column=0, sticky="e")
        self.end_current_entry = tk.Entry(root)
        self.end_current_entry.grid(row=1, column=1)

        tk.Label(root, text="Step Current (A):").grid(row=2, column=0, sticky="e")
        self.step_current_entry = tk.Entry(root)
        self.step_current_entry.grid(row=2, column=1)

        # Start button
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep)
        self.start_button.grid(row=3, column=0, columnspan=2, pady=10)

        # Matplotlib figure for plotting
        self.figure = plt.Figure(figsize=(7, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=4, column=0, columnspan=2, sticky="nsew")

        # Allow dynamic resizing
        root.grid_rowconfigure(4, weight=1)
        root.grid_columnconfigure(1, weight=1)

    def start_sweep(self):
        try:
            i_start = float(self.start_current_entry.get())
            i_end = float(self.end_current_entry.get())
            i_step = float(self.step_current_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers for all current fields.")
            return

        if i_step == 0:
            messagebox.showerror("Input Error", "Step current cannot be zero.")
            return

        self.start_button.config(state='disabled')

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save IV Data As"
        )
        if not file_path:
            self.start_button.config(state='normal')
            return

        # Connect to instrument
        rm = pyvisa.ResourceManager()
        try:
            load = rm.open_resource("USB0::0x2EC7::0x8800::802199042787070066::INSTR")
            load.write("MODE CC")
            load.write("INPUT ON")
        except Exception as e:
            messagebox.showerror("Connection Error", f"Could not connect to instrument:\n{e}")
            self.start_button.config(state='normal')
            return

        currents, voltages, powers = [], [], []
        current = i_start
        step = i_step if i_end >= i_start else -i_step

        while (step > 0 and current <= i_end) or (step < 0 and current >= i_end):
            try:
                load.write(f"CURR {current:.3f}")
                time.sleep(0.2)
                voltage = float(load.query("MEAS:VOLT?"))
                actual_current = float(load.query("MEAS:CURR?"))
                power = voltage * actual_current

                currents.append(actual_current)
                voltages.append(voltage)
                powers.append(power)

                print(f"I = {actual_current:.3f} A, V = {voltage:.3f} V, P = {power:.3f} W")
            except Exception as e:
                print(f"Measurement failed at {current:.3f} A: {e}")
                break

            current += step

        load.write("INPUT OFF")
        load.close()

        # Save to CSV
        try:
            with open(file_path, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["Current (A)", "Voltage (V)", "Power (W)"])
                for i, v, p in zip(currents, voltages, powers):
                    writer.writerow([i, v, p])
        except Exception as e:
            messagebox.showerror("File Error", f"Could not save file:\n{e}")
            self.start_button.config(state='normal')
            return

        # Plotting
        self.ax.clear()
        self.ax.plot(voltages, currents, 'b-o', label="I-V Curve")
        self.ax.set_xlabel("Voltage (V)")
        self.ax.set_ylabel("Current (A)", color='b')
        self.ax.tick_params(axis='y', labelcolor='b')
        self.ax.grid(True)

        # Power curve
        ax2 = self.ax.twinx()
        ax2.plot(voltages, powers, 'r--', label="Power Curve")
        ax2.set_ylabel("Power (W)", color='r')
        ax2.tick_params(axis='y', labelcolor='r')

        # Highlight Pmp
        if powers:
            pmp = max(powers)
            idx = powers.index(pmp)
            vmp = voltages[idx]
            imp = currents[idx]
            ax2.plot(vmp, pmp, 'ro')
            ax2.annotate(f"Pmp = {pmp:.2f} W", (vmp, pmp), textcoords="offset points", xytext=(0, 10), ha='center', color='r')

        self.figure.tight_layout()
        self.canvas.draw()
        self.start_button.config(state='normal')

        messagebox.showinfo("Sweep Complete", f"Sweep completed and data saved to:\n{file_path}")

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("850x750")
    root.resizable(True, True)
    app = IVAppCC(root)
    root.mainloop()
