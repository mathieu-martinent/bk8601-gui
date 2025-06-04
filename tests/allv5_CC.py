import tkinter as tk  # GUI library
from tkinter import messagebox, filedialog, ttk
import matplotlib.pyplot as plt  # Plotting library
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pyvisa  # VISA protocol for instrument communication
import time  # For delays
import csv  # For saving data to CSV
import os  # For file path manipulation
from datetime import datetime  # For timestamps


class IVAppCC:
    def __init__(self, root):
        self.root = root
        self.root.title("I-V Curve Measurement (CC/CV Mode)")

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

        # Operation mode selection
        tk.Label(root, text="Operation Mode:").grid(row=2, column=0, sticky="e")
        self.mode_var = tk.StringVar()
        self.mode_dropdown = ttk.Combobox(root, textvariable=self.mode_var, values=["CC", "CV"], state="readonly")
        self.mode_dropdown.grid(row=2, column=1, columnspan=2, sticky="ew")
        self.mode_dropdown.current(0)  # Default to CC mode
        self.mode_var.trace("w", self.update_labels)  # Update labels dynamically

        # Input fields for start, end, and step
        self.start_label = tk.Label(root, text="Start Current (A):")
        self.start_label.grid(row=3, column=0, sticky="e")
        self.start_current_entry = tk.Entry(root)
        self.start_current_entry.grid(row=3, column=1, columnspan=2, sticky="ew")

        self.end_label = tk.Label(root, text="End Current (A):")
        self.end_label.grid(row=4, column=0, sticky="e")
        self.end_current_entry = tk.Entry(root)
        self.end_current_entry.grid(row=4, column=1, columnspan=2, sticky="ew")

        self.step_label = tk.Label(root, text="Step Current (A):")
        self.step_label.grid(row=5, column=0, sticky="e")
        self.step_current_entry = tk.Entry(root)
        self.step_current_entry.grid(row=5, column=1, columnspan=2, sticky="ew")

        # Save options
        self.save_csv_var = tk.BooleanVar()
        self.save_csv_check = tk.Checkbutton(root, text="Save CSV data", variable=self.save_csv_var)
        self.save_csv_check.grid(row=6, column=0, columnspan=3, sticky="w", padx=5)

        self.save_png_var = tk.BooleanVar()
        self.save_png_check = tk.Checkbutton(root, text="Save plot as PNG", variable=self.save_png_var)
        self.save_png_check.grid(row=7, column=0, columnspan=3, sticky="w", padx=5)

        # Start button
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep)
        self.start_button.grid(row=8, column=0, columnspan=3, pady=10)

        # Progress bar
        self.progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')
        self.progress.grid(row=9, column=0, columnspan=3, sticky="ew", padx=5)

        # Plot area
        self.figure = plt.Figure(figsize=(7, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=10, column=0, columnspan=3, sticky="nsew")

        # Resizing support
        root.grid_rowconfigure(10, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Bind the Enter key to start_sweep
        self.root.bind('<Return>', lambda event: self.start_sweep())

    def update_labels(self, *args):
        """Update labels dynamically based on the selected operation mode."""
        selected_mode = self.mode_var.get()
        if selected_mode == "CC":
            self.start_label.config(text="Start Current (A):")
            self.end_label.config(text="End Current (A):")
            self.step_label.config(text="Step Current (A):")
        elif selected_mode == "CV":
            self.start_label.config(text="Start Voltage (V):")
            self.end_label.config(text="End Voltage (V):")
            self.step_label.config(text="Step Voltage (V):")

    def start_sweep(self):
        """Start the I-V sweep process."""
        try:
            i_start = float(self.start_current_entry.get())
            i_end = float(self.end_current_entry.get())
            i_step = float(self.step_current_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers for all fields.")
            return

        if i_step == 0:
            messagebox.showerror("Input Error", "Step value cannot be zero.")
            return

        if not self.instr_var.get():
            messagebox.showerror("Connection Error", "No instrument selected.")
            return

        self.start_button.config(state='disabled')

        instrument_address = self.instr_var.get()

        try:
            load = self.rm.open_resource(instrument_address)
            load.timeout = 5000  # 5 second timeout
            response = load.query("*IDN?")
            print(f"Instrument response: {response}")  # Debug print

            # Clear errors
            load.write("*CLS")

            # Query and set operation mode
            selected_mode = self.mode_var.get()
            current_mode = load.query("FUNC?").strip()
            print(f"Instrument is currently in {current_mode} mode.")  # Debug print

            # Map user-friendly mode names to SCPI commands
            mode_mapping = {"CC": "CURR", "CV": "VOLT"}
            if current_mode != mode_mapping[selected_mode]:
                print(f"Switching instrument mode to {selected_mode}.")
                load.write(f"FUNC {mode_mapping[selected_mode]}")

            load.write("INPUT ON")

            

            # Set sense mode
            sense_command = "REM:SENS ON" if self.sense_mode_var.get() == "4-Wire" else "REM:SENS OFF"
            load.write(sense_command)

            # Wait for the instrument to apply the settings
            time.sleep(0.5)

            # Query the actual sense mode
            actual_sense = load.query("REM:SENS?")
            print(f"Sense mode query response: {actual_sense.strip()}")  # Debug print

            # Initialize plot
            self.ax.clear()
            self.ax.set_xlabel("Voltage (V)" if selected_mode == "CV" else "Current (A)")
            self.ax.set_ylabel("Current (A)" if selected_mode == "CV" else "Voltage (V)", color='b')
            self.ax.tick_params(axis='y', labelcolor='b')
            self.ax.grid(True)
            ax2 = self.ax.twinx()
            ax2.set_ylabel("Power (W)", color='r')
            ax2.tick_params(axis='y', labelcolor='r')

            line_iv, = self.ax.plot([], [], 'b-o', label="I-V Curve")
            line_power, = ax2.plot([], [], 'r--', label="Power Curve")
            self.canvas.draw()

            # Sweep logic
            currents, voltages, powers = [], [], []
            current = i_start
            step = i_step if i_end >= i_start else -i_step
            total_steps = int(abs((i_end - i_start) / i_step)) + 1
            self.progress["maximum"] = total_steps
            self.progress["value"] = 0

            for count in range(total_steps):
                try:
                    if selected_mode == "CC":
                        load.write(f"CURR {current:.3f}")
                    elif selected_mode == "CV":
                        load.write(f"VOLT {current:.3f}")

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
                    print(f"Measurement failed at {current:.3f}: {e}")
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
                             (vmp, pmp), textcoords="offset points", xytext=(0, 10), ha='center', color='r')
                self.canvas.draw()
            else:
                pmp = vmp = imp = None

            # Generate a timestamped filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_filename = f"IV_Sweep_{timestamp}"
            output_dir = "output"
            os.makedirs(output_dir, exist_ok=True)

            # Save CSV file
            csv_path = os.path.join(output_dir, f"{base_filename}.csv")
            with open(csv_path, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["Current (A)", "Voltage (V)", "Power (W)"])
                for i, v, p in zip(currents, voltages, powers):
                    writer.writerow([i, v, p])
            print(f"Data saved to {csv_path}")

            # Save PNG file
            png_path = os.path.join(output_dir, f"{base_filename}.png")
            self.figure.savefig(png_path)
            print(f"Plot saved to {png_path}")

            self.start_button.config(state='normal')

            if pmp is not None:
                summary = f"Sweep completed successfully.\n\nPmp = {pmp:.2f} W\nVmp = {vmp:.2f} V\nImp = {imp:.2f} A"
            else:
                summary = "Sweep completed, but no power data was collected."
            messagebox.showinfo("Sweep Complete", summary)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")
            self.start_button.config(state='normal')


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("950x800")
    root.resizable(True, True)
    app = IVAppCC(root)
    root.mainloop()