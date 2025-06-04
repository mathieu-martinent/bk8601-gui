import tkinter as tk
from tkinter import messagebox, filedialog
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pyvisa
import time
import csv
import os
from tkinter import ttk
from datetime import datetime

class IVAppCC:
    def __init__(self, root):
        self.root = root
        self.root.title("I-V Curve Measurement (CC Mode)")

        # Current input fields
        tk.Label(root, text="Start Current (A):").grid(row=0, column=0, sticky="e")
        self.start_current_entry = tk.Entry(root)
        self.start_current_entry.grid(row=0, column=0, columnspan=2)

        tk.Label(root, text="End Current (A):").grid(row=1, column=0, sticky="e")
        self.end_current_entry = tk.Entry(root)
        self.end_current_entry.grid(row=1, column=0, columnspan=2)

        tk.Label(root, text="Step Current (A):").grid(row=2, column=0, sticky="e")
        self.step_current_entry = tk.Entry(root)
        self.step_current_entry.grid(row=2, column=0, columnspan=2)

        # Save options
        self.save_csv_var = tk.BooleanVar()
        self.save_csv_check = tk.Checkbutton(root, text="Save CSV data", variable=self.save_csv_var)
        self.save_csv_check.grid(row=3, column=0, columnspan=2, sticky="w", padx=5)

        self.save_png_var = tk.BooleanVar()
        self.save_png_check = tk.Checkbutton(root, text="Save plot as PNG", variable=self.save_png_var)
        self.save_png_check.grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=(0, 10))

        # Start button
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep)
        self.start_button.grid(row=5, column=0, columnspan=2, pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')
        self.progress.grid(row=6, column=0, columnspan=2, sticky="ew", padx=5, pady=(0, 10))

        # Plot area
        self.figure = plt.Figure(figsize=(7, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=7, column=0, columnspan=2, sticky="nsew")

        # Resizing support
        root.grid_rowconfigure(7, weight=1)
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

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = None
        if self.save_csv_var.get():
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv")],
                title="Save IV Data As",
                initialfile=f"IV_Data_{timestamp}.csv"
            )
            if not file_path:
                self.start_button.config(state='normal')
                return

        rm = pyvisa.ResourceManager()
        try:
            load = rm.open_resource("USB0::0x2EC7::0x8800::802199042787070066::INSTR")
            load.write("MODE CC")
            load.write("INPUT ON")
        except Exception as e:
            messagebox.showerror("Connection Error", f"Could not connect to instrument:\n{e}")
            self.start_button.config(state='normal')
            return

        self.ax.clear()
        self.ax.set_xlabel("Voltage (V)")
        self.ax.set_ylabel("Current (A)", color='b')
        self.ax.tick_params(axis='y', labelcolor='b')
        self.ax.grid(True)
        ax2 = self.ax.twinx()
        ax2.set_ylabel("Power (W)", color='r')
        ax2.tick_params(axis='y', labelcolor='r')

        line_iv, = self.ax.plot([], [], 'b-o', label="I-V Curve")
        line_power, = ax2.plot([], [], 'r--', label="Power Curve")
        self.canvas.draw()

        currents, voltages, powers = [], [], []
        current = i_start
        step = i_step if i_end >= i_start else -i_step
        total_steps = int(abs((i_end - i_start) / i_step)) + 1
        self.progress["maximum"] = total_steps
        self.progress["value"] = 0

        for count in range(total_steps):
            try:
                load.write(f"CURR {current:.3f}")
                time.sleep(0.2)
                voltage = float(load.query("MEAS:VOLT?"))
                actual_current = float(load.query("MEAS:CURR?"))
                power = voltage * actual_current

                currents.append(actual_current)
                voltages.append(voltage)
                powers.append(power)

                line_iv.set_data(voltages, currents)
                line_power.set_data(voltages, powers)

                self.ax.relim()
                self.ax.autoscale_view()
                ax2.relim()
                ax2.autoscale_view()

                self.canvas.draw()
                self.root.update_idletasks()
                self.progress["value"] = count + 1
            except Exception as e:
                print(f"Measurement failed at {current:.3f} A: {e}")
                break
            current += step

        load.write("INPUT OFF")
        load.close()

        if powers:
            pmp = max(powers)
            idx = powers.index(pmp)
            vmp = voltages[idx]
            imp = currents[idx]
            ax2.plot(vmp, pmp, 'ro')
            ax2.annotate(f"Pmp = {pmp:.2f} W\nVmp = {vmp:.2f} V\nImp = {imp:.2f} A",
                         (vmp, pmp), textcoords="offset points", xytext=(0, 10),
                         ha='center', color='r')
            self.canvas.draw()
        else:
            pmp = vmp = imp = None

        if file_path:
            try:
                with open(file_path, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(["Current (A)", "Voltage (V)", "Power (W)"])
                    for i, v, p in zip(currents, voltages, powers):
                        writer.writerow([i, v, p])
            except Exception as e:
                messagebox.showerror("File Error", f"Could not save CSV file:\n{e}")

        if self.save_png_var.get():
            if file_path:
                img_path = os.path.splitext(file_path)[0] + ".png"
            else:
                img_path = filedialog.asksaveasfilename(
                    defaultextension=".png",
                    filetypes=[("PNG files", "*.png")],
                    title="Save Plot As",
                    initialfile=f"IV_Plot_{timestamp}.png"
                )
            if img_path:
                try:
                    self.figure.savefig(img_path)
                except Exception as e:
                    messagebox.showerror("Save Error", f"Could not save PNG file:\n{e}")

        self.start_button.config(state='normal')

        if pmp is not None:
            summary = f"Sweep completed successfully.\n\nPmp = {pmp:.2f} W\nVmp = {vmp:.2f} V\nImp = {imp:.2f} A"
        else:
            summary = "Sweep completed, but no power data was collected."
        messagebox.showinfo("Sweep Complete", summary)

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("900x750")
    root.resizable(True, True)
    app = IVAppCC(root)
    root.mainloop()
