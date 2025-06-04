import tkinter as tk # used for creating the graphical user interface (GUI) for the I-V curve measurement application
from tkinter import messagebox, filedialog, ttk                     
import matplotlib.pyplot as plt # used for plotting the I-V and power curve                                               
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg     
import pyvisa # used for communicating with the instrument via VISA protocol                  
import time # used for adding delays during communication with the instrument                 
import csv # used for saving the measurement data to a CSV file                         
import os # used for file path manipulation                             
from datetime import datetime # used for generating timestamps for file names                                

class IVAppCC:

    # Initializes the application and sets the window title 
    def __init__(self, root):
        self.root = root
        self.root.title("I-V Curve Measurement (CC Mode)")

        # VISA instrument selection
        tk.Label(root, text="Select Instrument:").grid(row=0, column=0, sticky="e")
        self.rm = pyvisa.ResourceManager()
        self.instr_list = self.rm.list_resources()
        
        self.instr_var = tk.StringVar()
        self.instr_dropdown = ttk.Combobox(root, textvariable=self.instr_var, values=self.instr_list, state="readonly")
        self.instr_dropdown.grid(row=0, column=1, columnspan=2, sticky="ew")
        if self.instr_list:
            self.instr_dropdown.current(0)

        # Sense mode selection
        tk.Label(root, text="Sense Mode:").grid(row=1, column=0, sticky="e")
        self.sense_mode_var = tk.StringVar()
        self.sense_mode_dropdown = ttk.Combobox(root, textvariable=self.sense_mode_var, values=["2-Wire", "4-Wire"], state="readonly")
        self.sense_mode_dropdown.grid(row=1, column=1, columnspan=2, sticky="ew")
        self.sense_mode_dropdown.current(0)

        # Current input fields
        tk.Label(root, text="Start Current (A):").grid(row=2, column=0, sticky="e")
        self.start_current_entry = tk.Entry(root)
        self.start_current_entry.grid(row=2, column=1, columnspan=2, sticky="ew")

        tk.Label(root, text="End Current (A):").grid(row=3, column=0, sticky="e")
        self.end_current_entry = tk.Entry(root)
        self.end_current_entry.grid(row=3, column=1, columnspan=2, sticky="ew")

        tk.Label(root, text="Step Current (A):").grid(row=4, column=0, sticky="e")
        self.step_current_entry = tk.Entry(root)
        self.step_current_entry.grid(row=4, column=1, columnspan=2, sticky="ew")

        # Save options
        self.save_csv_var = tk.BooleanVar()
        self.save_csv_check = tk.Checkbutton(root, text="Save CSV data", variable=self.save_csv_var)
        self.save_csv_check.grid(row=5, column=0, columnspan=3, sticky="w", padx=5)

        self.save_png_var = tk.BooleanVar()
        self.save_png_check = tk.Checkbutton(root, text="Save plot as PNG", variable=self.save_png_var)
        self.save_png_check.grid(row=6, column=0, columnspan=3, sticky="w", padx=5)

        # Start button
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep)
        self.start_button.grid(row=7, column=0, columnspan=3, pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')
        self.progress.grid(row=8, column=0, columnspan=3, sticky="ew", padx=5)

        # Plot area
        self.figure = plt.Figure(figsize=(7, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=9, column=0, columnspan=3, sticky="nsew")

        # Resizing support
        root.grid_rowconfigure(9, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Bind the Enter key to start_sweep
        self.root.bind('<Return>', lambda event: self.start_sweep())

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

        if not self.instr_var.get():
            messagebox.showerror("Connection Error", "No instrument selected.")
            return

        self.start_button.config(state='disabled')

        instrument_address = self.instr_var.get()
      
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

        try:
            load = self.rm.open_resource(instrument_address)
            load.timeout = 5000  # 5 second timeout
            response = load.query("*IDN?")
            time.sleep(0.1)  # Add a 100ms delay
            print(f"Instrument response: {response}")  # Debug print

            # Clear errors and set mode to CC
            load.write("*CLS")
            load.write("MODE CC")
            load.write("INPUT ON")

            # Set sense mode
            sense_command = "REM:SENS ON" if self.sense_mode_var.get() == "4-Wire" else "REM:SENS OFF"
            print(f"Sending sense mode command: {sense_command}")  # Debug print
            load.write(sense_command)

            # Wait for the instrument to apply the setting
            time.sleep(0.5)

             # Query the actual sense mode
            actual_sense = load.query("REM:SENS?")
            print(f"Sense mode query response: {actual_sense.strip()}")  # Debug print
            if actual_sense.strip() == '1':
                print("Instrument is in 4-Wire mode.")
            else:
                print("Instrument is in 2-Wire mode.")
        
        
        except Exception as e:
            print(f"Error communicating with instrument: {e}")  # Debug print
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
    root.geometry("950x800")
    root.resizable(True, True)
    app = IVAppCC(root)
    root.mainloop()
