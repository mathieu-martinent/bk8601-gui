import tkinter as tk
from tkinter import messagebox, filedialog, ttk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pyvisa
import time
import csv
import os
from datetime import datetime
import math
import threading
import json
import pandas as pd  # Add this import

class IVAppCC:
    def __init__(self, root):
        self.root = root
        self.root.title("I-V Curve Measurement (CC/CV Mode)")

        # VISA instrument selection dropdown
        tk.Label(root, text="Select Instrument:").grid(row=0, column=0, sticky="e")
        try:
            self.rm = pyvisa.ResourceManager()
            real_instr = list(self.rm.list_resources())
        except Exception:
            self.rm = None
            real_instr = []
        self.instr_list = ["Simulated Instrument"] + real_instr
        self.instr_var = tk.StringVar()
        self.instr_dropdown = ttk.Combobox(root, textvariable=self.instr_var, values=self.instr_list, state="readonly")
        self.instr_dropdown.grid(row=0, column=1, columnspan=2, sticky="ew")
        self.instr_dropdown.current(0)

        # Sense mode selection (2-Wire or 4-Wire)
        tk.Label(root, text="Sense Mode:").grid(row=1, column=0, sticky="e")
        self.sense_mode_var = tk.StringVar()
        self.sense_mode_dropdown = ttk.Combobox(root, textvariable=self.sense_mode_var, values=["2-Wire", "4-Wire"], state="readonly")
        self.sense_mode_dropdown.grid(row=1, column=1, columnspan=2, sticky="ew")
        self.sense_mode_dropdown.current(0)

        # Operation mode selection (CC or CV)
        tk.Label(root, text="Operation Mode:").grid(row=2, column=0, sticky="e")
        self.mode_var = tk.StringVar()
        self.mode_dropdown = ttk.Combobox(root, textvariable=self.mode_var, values=["CC", "CV"], state="readonly")
        self.mode_dropdown.grid(row=2, column=1, columnspan=2, sticky="ew")
        self.mode_dropdown.current(0)
        self.mode_var.trace("w", self.update_labels)

        # Input fields for start, end, and step values
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

        # Protection settings for voltage and current
        tk.Label(root, text="Voltage Limit (V):").grid(row=6, column=0, sticky="e")
        self.voltage_limit_entry = tk.Entry(root)
        self.voltage_limit_entry.grid(row=6, column=1, columnspan=2, sticky="ew")

        tk.Label(root, text="Current Limit (A):").grid(row=7, column=0, sticky="e")
        self.current_limit_entry = tk.Entry(root)
        self.current_limit_entry.grid(row=7, column=1, columnspan=2, sticky="ew")

        # Sleep time between steps
        tk.Label(root, text="Step Delay (s):").grid(row=8, column=0, sticky="e")
        self.sleep_time_entry = tk.Entry(root)
        self.sleep_time_entry.grid(row=8, column=1, columnspan=2, sticky="ew")
        self.sleep_time_entry.insert(0, "0.5")

        # Save options for CSV and PNG files
        self.save_csv_var = tk.BooleanVar()
        self.save_csv_check = tk.Checkbutton(root, text="Save CSV data", variable=self.save_csv_var)
        self.save_csv_check.grid(row=9, column=0, columnspan=3, sticky="w", padx=5)

        self.save_png_var = tk.BooleanVar()
        self.save_png_check = tk.Checkbutton(root, text="Save plot as PNG", variable=self.save_png_var)
        self.save_png_check.grid(row=10, column=0, columnspan=3, sticky="w", padx=5)

        # Output directory at project root (not in src)
        self.output_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'output'))
        os.makedirs(self.output_dir, exist_ok=True)
        self.folder_label = tk.Label(root, text=f"Output directory: {self.output_dir}")
        self.folder_label.grid(row=14, column=0, columnspan=2, sticky="w", padx=5, pady=(5, 0))
        self.folder_button = tk.Button(root, text="Change...", command=self.choose_output_dir)
        self.folder_button.grid(row=14, column=2, sticky="e", padx=5, pady=(5, 0))

        # Start and Stop buttons
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep_thread)
        self.start_button.grid(row=11, column=0, columnspan=2, pady=10)

        self.stop_button = tk.Button(root, text="Stop", command=self.request_stop, state="normal")
        self.stop_button.grid(row=11, column=2, pady=10)

        # Add comparison button after the stop button
        self.compare_button = tk.Button(root, text="Compare Curves", command=self.open_comparison_window)
        self.compare_button.grid(row=11, column=3, pady=10)

        # Progress bar to show sweep progress
        self.progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')
        self.progress.grid(row=12, column=0, columnspan=3, sticky="ew", padx=5)

        # Plot area for displaying the I-V curve
        self.figure = plt.Figure(figsize=(7, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=13, column=0, columnspan=3, sticky="nsew")

        # Configure resizing behavior for the plot area
        root.grid_rowconfigure(13, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Flags for sweep state
        self.stop_requested = False
        self.sweep_running = False

        # Bind Enter and Escape keys for quick actions
        self.root.bind('<Return>', self.on_enter)
        self.root.bind('<Escape>', self.on_escape)

        # Load last used settings from file
        self.load_settings()

    def choose_output_dir(self):
        """Open a dialog to choose a new output directory."""
        new_dir = filedialog.askdirectory(initialdir=os.getcwd(), title="Choose output directory")
        if new_dir:
            self.output_dir = new_dir
            self.folder_label.config(text=f"Output directory: {self.output_dir}")

    def on_enter(self, event):
        """Start sweep when Enter is pressed, if not already running."""
        if not self.sweep_running:
            self.start_sweep_thread()

    def on_escape(self, event):
        """Request stop when Escape is pressed, if sweep is running."""
        if self.sweep_running:
            self.request_stop()

    def start_sweep_thread(self):
        """Start the sweep in a separate thread to avoid freezing the GUI."""
        if self.sweep_running:
            return  # Prevent double start
        self.sweep_running = True
        thread = threading.Thread(target=self.start_sweep)
        thread.daemon = True
        thread.start()

    def request_stop(self):
        """Set the stop flag to request sweep interruption."""
        self.stop_requested = True

    def update_labels(self, *args):
        """Update the labels of the input fields depending on the selected mode."""
        selected_mode = self.mode_var.get()
        if selected_mode == "CC":
            self.start_label.config(text="Start Current (A):")
            self.end_label.config(text="End Current (A):")
            self.step_label.config(text="Step Current (A):")
        else:
            self.start_label.config(text="Start Voltage (V):")
            self.end_label.config(text="End Voltage (V):")
            self.step_label.config(text="Step Voltage (V):")

    def save_settings(self):
        """Save the current GUI settings to a JSON file."""
        settings = {
            "start": self.start_current_entry.get(),
            "end": self.end_current_entry.get(),
            "step": self.step_current_entry.get(),
            "voltage_limit": self.voltage_limit_entry.get(),
            "current_limit": self.current_limit_entry.get(),
            "sleep_time": self.sleep_time_entry.get(),
            "mode": self.mode_var.get(),
            "sense": self.sense_mode_var.get(),
            "instr": self.instr_var.get(),
            "save_csv": self.save_csv_var.get(),
            "save_png": self.save_png_var.get(),
            "output_dir": self.output_dir,
        }
        with open("last_settings.json", "w") as f:
            json.dump(settings, f)

    def load_settings(self):
        """Load the last used GUI settings from a JSON file."""
        if os.path.exists("last_settings.json"):
            with open("last_settings.json", "r") as f:
                settings = json.load(f)
            self.start_current_entry.delete(0, tk.END)
            self.start_current_entry.insert(0, settings.get("start", ""))
            self.end_current_entry.delete(0, tk.END)
            self.end_current_entry.insert(0, settings.get("end", ""))
            self.step_current_entry.delete(0, tk.END)
            self.step_current_entry.insert(0, settings.get("step", ""))
            self.voltage_limit_entry.delete(0, tk.END)
            self.voltage_limit_entry.insert(0, settings.get("voltage_limit", ""))
            self.current_limit_entry.delete(0, tk.END)
            self.current_limit_entry.insert(0, settings.get("currentLimit", ""))
            self.sleep_time_entry.delete(0, tk.END)
            self.sleep_time_entry.insert(0, settings.get("sleep_time", ""))
            self.mode_var.set(settings.get("mode", "CC"))
            self.sense_mode_var.set(settings.get("sense", "2-Wire"))
            if settings.get("instr") in self.instr_list:
                self.instr_var.set(settings.get("instr"))
            self.save_csv_var.set(settings.get("save_csv", False))
            self.save_png_var.set(settings.get("save_png", False))
            if os.path.isdir(settings.get("output_dir", "")):
                self.output_dir = settings["output_dir"]
                self.folder_label.config(text=f"Output directory: {self.output_dir}")

    def start_sweep(self):
        """Main function to perform the I-V sweep and handle all measurements, plotting, and saving."""
        self.stop_requested = False  # Reset stop flag at each sweep start
        try:
            # Parse user input values from GUI
            i_start = float(self.start_current_entry.get())
            i_end = float(self.end_current_entry.get())
            i_step = float(self.step_current_entry.get())
            voltage_limit = self.voltage_limit_entry.get()
            current_limit = self.current_limit_entry.get()
            voltage_limit = float(voltage_limit) if voltage_limit else None
            current_limit = float(current_limit) if current_limit else None
            sleep_time = float(self.sleep_time_entry.get())
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return

        selected_mode = self.mode_var.get()
        # --- Mandatory safety check ---
        if selected_mode == "CC" and voltage_limit is None:
            messagebox.showerror("Safety", "In CC mode, the Voltage Limit is mandatory to protect the solar cell.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return
        if selected_mode == "CV" and current_limit is None:
            messagebox.showerror("Safety", "In CV mode, the Current Limit is mandatory to protect the solar cell.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return
        # --- End safety check ---

        if i_step == 0:
            messagebox.showerror("Input Error", "Step value cannot be zero.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return

        if not self.instr_var.get():
            messagebox.showerror("Connection Error", "No instrument selected.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return

        self.start_button.config(state='disabled')
        instrument_address = self.instr_var.get()

        try:
            # Open instrument connection (real or simulated)
            if instrument_address == "Simulated Instrument":
                load = self.create_simulated_instrument()
            else:
                load = self.rm.open_resource(instrument_address)
                load.timeout = 5000
                load.write("*RST")   # Reset instrument
                load.write("*CLS")   # Clear status

            # Set instrument mode (CC or CV)
            selected_mode = self.mode_var.get()
            mode_mapping = {"CC": "CURR", "CV": "VOLT"}
            if load.query("FUNC?").strip() != mode_mapping[selected_mode]:
                load.write(f"FUNC {mode_mapping[selected_mode]}")

            # Configure protection limits
            if voltage_limit is not None:
                load.write("VOLT:PROT:STAT ON")
                load.write(f"VOLT:PROT {voltage_limit}")
            else:
                load.write("VOLT:PROT:STAT OFF")

            if current_limit is not None:
                load.write("CURR:PROT:STAT ON")
                load.write(f"CURR:PROT {current_limit}")
            else:
                load.write("CURR:PROT:STAT OFF")

            # Set sense mode (2-Wire or 4-Wire)
            sense_command = "REM:SENS ON" if self.sense_mode_var.get() == "4-Wire" else "REM:SENS OFF"
            load.write(sense_command)
            time.sleep(0.5)

            # Enable the input only after all configuration is done
            load.write("INPUT ON")

            # Prepare plot area
            self.ax.clear()
            if hasattr(self, 'ax2'):
                self.figure.delaxes(self.ax2)
            self.ax2 = self.ax.twinx()

            # Set axis labels
            self.ax.set_xlabel("Voltage (V)")
            self.ax.set_ylabel("Current (A)", color='b')
            self.ax2.set_ylabel("Power (W)", color='r')

            # Remove previous plot lines and annotations if they exist
            for attr in ['line_iv', 'line_power', 'pmp_annotation', 'pmp_point', 'vmp_annotation', 'vmp_point', 'summary_annotation']:
                if hasattr(self, attr):
                    try:
                        getattr(self, attr).remove()
                    except Exception:
                        pass
                    delattr(self, attr)         

            self.ax.set_xlabel("Voltage (V)")
            self.ax.set_ylabel("Current (A)", color='b')
            self.ax.tick_params(axis='y', labelcolor='b')
            self.ax.grid(True)
            self.ax2.yaxis.set_label_position("right")
            self.ax2.yaxis.tick_right()
            self.ax2.set_ylabel("Power (W)", color='r')
            self.ax2.tick_params(axis='y', labelcolor='r')
            self.canvas.draw()

            # Prepare data lists for sweep results
            currents, voltages, powers = [], [], []

            # Configure sweep parameters and setpoint command depending on mode
            if selected_mode == "CC":
                sweep_start = i_start
                sweep_end = i_end
                sweep_step = i_step if i_end >= i_start else -i_step
                setpoint_cmd = lambda v: load.write(f"CURR {v:.3f}")
            else:
                sweep_start = i_start
                sweep_end = i_end
                sweep_step = i_step if i_end >= i_start else -i_step
                setpoint_cmd = lambda v: load.write(f"VOLT {v:.3f}")

            value = sweep_start
            total_steps = int(abs((sweep_end - sweep_start) / sweep_step)) + 1
            self.progress["maximum"] = total_steps
            self.progress["value"] = 0
            print(f"total_steps = {total_steps}, sweep_start = {sweep_start}, sweep_end = {sweep_end}, sweep_step = {sweep_step}")  # Debugging line

            # Force starting setpoint and wait before sweep
            setpoint_cmd(sweep_start)
            time.sleep(sleep_time)

            # Enable input just before starting the sweep
            load.write("INPUT ON")
            time.sleep(0.2)

            for count in range(total_steps):
                if self.stop_requested:
                    messagebox.showinfo("Sweep Stopped", "Sweep was stopped by the user.")
                    break
                try:
                    setpoint_cmd(value)         # Set current or voltage depending on mode
                    time.sleep(sleep_time)
                    voltage = float(load.query("MEAS:VOLT?"))
                    actual_current = float(load.query("MEAS:CURR?"))
                    power = voltage * actual_current

                    # Protection checks
                    if voltage_limit is not None and voltage > voltage_limit:
                        raise Exception("Voltage exceeded protection limit.")
                    if current_limit is not None and actual_current > current_limit:
                        raise Exception("Current exceeded protection limit.")

                    print(f"Protection check: V={voltage} (limit {voltage_limit}), I={actual_current} (limit {current_limit})")
                    print(f"Setpoint: {value:.3f} V, Measured: {voltage:.3f} V, {actual_current:.3f} A")  # Debug print

                    # Only append new point if it is different from the last one
                    EPS = 1e-4
                    if len(currents) == 0 or abs(actual_current - currents[-1]) > EPS or abs(voltage - voltages[-1]) > EPS:
                        currents.append(actual_current)
                        voltages.append(voltage)
                        powers.append(power)

                    # Update I-V curve plot
                    if hasattr(self, 'line_iv'):
                        self.line_iv.remove()
                        del self.line_iv
                    self.line_iv, = self.ax.plot(voltages, currents, label="I-V Curve", color='blue')

                    # Update P-V curve plot
                    if hasattr(self, 'line_power'):
                        self.line_power.remove()
                        del self.line_power
                    self.line_power, = self.ax2.plot(voltages, powers, label="P-V Curve", color='red')

                    self.ax.relim()
                    self.ax.autoscale_view()
                    self.ax2.relim()
                    self.ax2.autoscale_view()

                    # Adjust X axis limits: always start at 0 V, end just above max measured voltage
                    if voltages:
                        v_max = max(voltages)
                        self.ax.set_xlim(left=0, right=v_max * 1.0105)
                    else:
                        self.ax.set_xlim(left=0)

                    self.canvas.draw()
                    self.root.update_idletasks()
                    self.progress["value"] = count + 1
                except Exception as e:
                    print(f"Exception in sweep loop: {e}")  # Print exception for debugging
                    messagebox.showwarning("Protection Triggered", f"Sweep stopped: {e}")
                    break
                value += sweep_step

            # Turn off instrument input and close connection
            load.write("INPUT OFF")
            load.close()

            # Final plot update after sweep
            if voltages and currents and hasattr(self, 'line_iv'):
                self.line_iv.set_data(voltages, currents)
            if voltages and powers and hasattr(self, 'line_power'):
                self.line_power.set_data(voltages, powers)
            self.ax.set_xlabel("Voltage (V)")
            self.ax.set_ylabel("Current (A)", color='b')
            self.ax2.set_ylabel("Power (W)", color='r')
            self.ax.relim()
            self.ax.autoscale_view()
            self.ax2.relim()
            self.ax2.autoscale_view()
            self.canvas.draw()

            # Annotate summary results on the plot
            if powers:
                pmp = max(powers)
                idx = powers.index(pmp)
                vmp = voltages[idx]
                imp = currents[idx]
                summary_text = f"Pmp = {pmp:.2f} W   Vmp = {vmp:.2f} V   Imp = {imp:.2f} A"
            else:
                pmp = vmp = imp = None
                summary_text = "Sweep completed with no power data."

            if hasattr(self, 'summary_annotation'):
                try:
                    self.summary_annotation.remove()
                except Exception:
                    pass
                del self.summary_annotation

            self.summary_annotation = self.ax.annotate(
                summary_text,
                xy=(0.5, 1.08), xycoords='axes fraction',
                ha='center', va='bottom',
                fontsize=14, color='purple',
                bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="purple", lw=2)
            )
            self.canvas.draw()

            # Prepare file and folder names for saving results
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            selected_mode = self.mode_var.get()
            sense_mode = self.sense_mode_var.get()
            base_filename = f"IV_Sweep_{selected_mode}_{sense_mode}_{timestamp}"

            unit = "A" if selected_mode == "CC" else "V"
            params = [
                ("Mode", selected_mode),
                ("Sense", sense_mode),
                ("Start (A)" if selected_mode == "CC" else "Start (V)", i_start),
                ("End (A)" if selected_mode == "CC" else "End (V)", i_end),
                ("Step (A)" if selected_mode == "CC" else "Step (V)", i_step),
                ("Voltage Limit (V)", voltage_limit),
                ("Current Limit (A)", current_limit),
                ("Step Delay (s)", sleep_time),
                ("Instrument", instrument_address),
            ]

            # Create a subfolder for today's date if it doesn't exist
            today_str = datetime.now().strftime("%Y-%m-%d")
            day_output_dir = os.path.join(self.output_dir, today_str)
            os.makedirs(day_output_dir, exist_ok=True)

            # Save CSV data if requested
            if self.save_csv_var.get():
                csv_path = os.path.join(day_output_dir, f"{base_filename}.csv")
                with open(csv_path, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(["Current (A)", "Voltage (V)", "Power (W)"])
                    for i in range(len(currents)):
                        writer.writerow([currents[i], voltages[i], powers[i]])
                    writer.writerow([])
                    writer.writerow(["Parameter", "Value"])
                    for param, value in params:
                        writer.writerow([param, value])
                print(f"Data saved to {csv_path}")

            # Highlight the Pmp point (maximum power) on the Power curve (red axis)
            if powers and voltages and currents:
                pmp = max(powers)
                idx = powers.index(pmp)
                vmp = voltages[idx]
                imp = currents[idx]
                # Plot a big red dot at the maximum power point on the P-V curve
                self.ax2.plot(vmp, pmp, 'ro', markersize=12, label="Pmp")
                # Add a label with an arrow pointing to the point
                self.ax2.annotate(
                    "Pmp",
                    xy=(vmp, pmp),
                    xytext=(20, 20),
                    textcoords='offset points',
                    fontsize=12,
                    color='red',
                    ha='left',
                    va='bottom',
                    arrowprops=dict(arrowstyle="->", color='red', lw=2),
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="red", lw=1)
                )
                # Force redraw so it's visible even before resizing
                self.canvas.draw()

            # Save PNG plot if requested
            if self.save_png_var.get():
                png_path = os.path.join(day_output_dir, f"{base_filename}.png")
                self.figure.savefig(png_path)
                print(f"Plot saved to {png_path}")

            # Show summary message box
            message = f"Sweep completed.\nPmp = {pmp:.2f} W\nVmp = {vmp:.2f} V\nImp = {imp:.2f} A" if pmp else "Sweep completed with no power data."
            messagebox.showinfo("Sweep Complete", message)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")

        finally:
            # Save settings after sweep completion
            self.save_settings()
            self.sweep_running = False
            self.start_button.config(state='normal')

    def create_simulated_instrument(self):
        """Create a simulated instrument for testing without hardware."""
        class SimulatedInstrument:
            def __init__(self):
                self.state = {
                    "FUNC": "CURR",
                    "current": 0.0,
                    "voltage": 0.0,
                    "VOLT_PROT_ON": False,
                    "VOLT_PROT": None,
                    "CURR_PROT_ON": False,
                    "CURR_PROT": None,
                }

            def write(self, command):
                # Simulate instrument command handling
                if "FUNC" in command:
                    self.state["FUNC"] = command.split()[-1]
                elif "CURR" in command and "CURR:PROT" not in command:
                    self.state["current"] = float(command.split()[-1])
                elif "VOLT" in command and "VOLT:PROT" not in command:
                    self.state["voltage"] = float(command.split()[-1])
                elif "VOLT:PROT:STAT ON" in command:
                    self.state["VOLT_PROT_ON"] = True
                elif "VOLT:PROT:STAT OFF" in command:
                    self.state["VOLT_PROT_ON"] = False
                elif "VOLT:PROT" in command:
                    try:
                        self.state["VOLT_PROT"] = float(command.split()[-1])
                    except Exception:
                        pass
                elif "CURR:PROT:STAT ON" in command:
                    self.state["CURR_PROT_ON"] = True
                elif "CURR:PROT:STAT OFF" in command:
                    self.state["CURR_PROT_ON"] = False
                elif "CURR:PROT" in command:
                    try:
                        self.state["CURR_PROT"] = float(command.split()[-1])
                    except Exception:
                        pass

            def query(self, command):
                # Simulate instrument measurement queries
                Isc = 5.0
                Voc = 25
                n = 1.5
                Vt = 0.7

                if "MEAS:VOLT?" in command:
                    # In CC mode, given current, return voltage
                    if self.state["FUNC"] == "CURR":
                        I = self.state["current"]
                        # Calculate V for given I (numerically invert the diode equation)
                        V = Voc + n * Vt * math.log(1 - I / Isc) if I < Isc else 0
                        if self.state["VOLT_PROT_ON"] and self.state["VOLT_PROT"] is not None and V > self.state["VOLT_PROT"]:
                            return str(self.state["VOLT_PROT"] + 5)
                        return str(max(V, 0))
                    return str(self.state["voltage"])
                elif "MEAS:CURR?" in command:
                    # In CV mode, given voltage, return current
                    if self.state["FUNC"] == "VOLT":
                        V = self.state["voltage"]
                        I = Isc * (1 - math.exp((V - Voc) / (n * Vt)))
                        if I < 0:
                            I = 0
                        if self.state["CURR_PROT_ON"] and self.state["CURR_PROT"] is not None and I > self.state["CURR_PROT"]:
                            return str(self.state["CURR_PROT"] + 5)
                        return str(I)
                    return str(self.state["current"])
                elif "FUNC?" in command:
                    return self.state.get("FUNC", "CURR")
                elif "STAT:QUES:COND?" in command:
                    return "0"
                return "0"

            def close(self):
                pass

        return SimulatedInstrument()

    def open_comparison_window(self):
        """Open a new window for comparing multiple I-V curves."""
        comparison_window = tk.Toplevel(self.root)
        comparison_window.title("I-V Curve Comparison")
        comparison_window.geometry("1200x800")
        
        # Create comparison app instance
        ComparisonApp(comparison_window, self.output_dir)

class ComparisonApp:
    def __init__(self, root, default_output_dir):
        self.root = root
        self.output_dir = default_output_dir
        self.loaded_curves = []  # List to store loaded curve data
        
        # Create main frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Control panel on the left
        control_frame = tk.Frame(main_frame)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # File selection area
        tk.Label(control_frame, text="Load I-V Data Files:", font=("Arial", 12, "bold")).pack(anchor="w")
        
        # Quick access to recent measurements
        tk.Button(control_frame, text="Browse Recent Measurements", 
                 command=self.browse_recent_measurements, bg="lightgreen").pack(pady=5)
        
        # Listbox to show loaded files
        self.file_listbox = tk.Listbox(control_frame, height=8, width=40)
        self.file_listbox.pack(pady=(5, 0))
        
        # Buttons for file operations
        button_frame = tk.Frame(control_frame)
        button_frame.pack(pady=5)
        
        tk.Button(button_frame, text="Add CSV File", command=self.add_csv_file).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(button_frame, text="Remove Selected", command=self.remove_selected_file).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(button_frame, text="Clear All", command=self.clear_all_files).pack(side=tk.LEFT)
        
        # Plot options
        tk.Label(control_frame, text="Plot Options:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(20, 5))
        
        # Checkboxes for what to plot
        self.show_iv_var = tk.BooleanVar(value=True)
        self.show_pv_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(control_frame, text="Show I-V Curves", variable=self.show_iv_var, command=self.update_plot).pack(anchor="w")
        tk.Checkbutton(control_frame, text="Show P-V Curves", variable=self.show_pv_var, command=self.update_plot).pack(anchor="w")
        
        # Plot button
        tk.Button(control_frame, text="Update Plot", command=self.update_plot, 
                 bg="lightblue", font=("Arial", 10, "bold")).pack(pady=10)
        
        # Export button
        tk.Button(control_frame, text="Export Comparison", command=self.export_comparison).pack(pady=(0, 10))
        
        # Statistics area
        tk.Label(control_frame, text="Statistics:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(20, 5))
        self.stats_text = tk.Text(control_frame, height=10, width=40, font=("Courier", 9))
        self.stats_text.pack()
        
        # Plot area on the right
        plot_frame = tk.Frame(main_frame)
        plot_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Create matplotlib figure
        self.figure = plt.Figure(figsize=(10, 8), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Initialize empty plot
        self.ax1 = self.figure.add_subplot(211)  # I-V curves
        self.ax2 = self.figure.add_subplot(212)  # P-V curves
        self.figure.tight_layout()
        self.canvas.draw()
    
    def browse_recent_measurements(self):
        """Browse and select from recent measurements in the output folder."""
        if not os.path.exists(self.output_dir):
            messagebox.showwarning("Warning", f"Output directory not found: {self.output_dir}")
            return
        
        # Create a selection window
        selection_window = tk.Toplevel(self.root)
        selection_window.title("Select Recent Measurements")
        selection_window.geometry("600x400")
        
        # Get all CSV files from output folder (organized by date)
        csv_files = []
        for root, dirs, files in os.walk(self.output_dir):
            for file in files:
                if file.endswith('.csv'):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, self.output_dir)
                    csv_files.append((rel_path, full_path))
        
        # Sort by modification time (newest first)
        csv_files.sort(key=lambda x: os.path.getmtime(x[1]), reverse=True)
        
        tk.Label(selection_window, text="Recent Measurements (newest first):", 
                font=("Arial", 12, "bold")).pack(pady=5)
        
        # Create listbox with scrollbar
        list_frame = tk.Frame(selection_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode=tk.MULTIPLE)
        file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=file_listbox.yview)
        
        # Populate listbox
        for rel_path, full_path in csv_files:
            file_listbox.insert(tk.END, rel_path)
        
        # Buttons
        button_frame = tk.Frame(selection_window)
        button_frame.pack(pady=10)
        
        def load_selected():
            selected_indices = file_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("Warning", "Please select at least one file.")
                return
            
            for idx in selected_indices:
                _, full_path = csv_files[idx]
                self.load_csv_file(full_path)
            
            selection_window.destroy()
        
        tk.Button(button_frame, text="Load Selected", command=load_selected, 
                 bg="lightgreen").pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=selection_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def add_csv_file(self):
        """Add a CSV file to the comparison."""
        file_path = filedialog.askopenfilename(
            title="Select I-V Data CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=self.output_dir
        )
        
        if file_path:
            self.load_csv_file(file_path)
    
    def load_csv_file(self, file_path):
        """Load a single CSV file."""
        try:
            # Try to read the CSV file
            df = pd.read_csv(file_path)
            
            # Check if required columns exist
            required_cols = ["Current (A)", "Voltage (V)", "Power (W)"]
            if not all(col in df.columns for col in required_cols):
                messagebox.showerror("Error", f"CSV file must contain columns: {', '.join(required_cols)}")
                return
            
            # Clean the data: remove non-numeric rows and convert to numeric
            # First, drop any rows with NaN values
            df = df.dropna()
            
            # Convert to numeric, coercing errors to NaN
            df["Current (A)"] = pd.to_numeric(df["Current (A)"], errors='coerce')
            df["Voltage (V)"] = pd.to_numeric(df["Voltage (V)"], errors='coerce')
            df["Power (W)"] = pd.to_numeric(df["Power (W)"], errors='coerce')
            
            # Drop rows that couldn't be converted to numbers
            df = df.dropna()
            
            # Additional check: remove rows where all values are zero (header repetitions)
            df = df[(df["Current (A)"] != 0) | (df["Voltage (V)"] != 0) | (df["Power (W)"] != 0)]
            
            if df.empty:
                messagebox.showerror("Error", "No valid numeric data found in CSV file")
                return
            
            # Extract metadata from filename if possible
            filename = os.path.basename(file_path)
            
            # Try to extract mode and sense from filename
            mode = "Unknown"
            sense = "Unknown"
            
            if "_CC_" in filename:
                mode = "CC"
            elif "_CV_" in filename:
                mode = "CV"
            
            if "_4-Wire_" in filename:
                sense = "4-Wire"
            elif "_2-Wire_" in filename:
                sense = "2-Wire"
            
            # If not found in filename, try to guess from data
            if mode == "Unknown":
                # Simple heuristic: if voltage range is larger, probably CV mode
                voltage_range = df["Voltage (V)"].max() - df["Voltage (V)"].min()
                current_range = df["Current (A)"].max() - df["Current (A)"].min()
                mode = "CV" if voltage_range > current_range else "CC"
            
            # Store the curve data
            curve_data = {
                'file_path': file_path,
                'filename': filename,
                'mode': mode,
                'sense': sense,
                'current': df["Current (A)"].values,
                'voltage': df["Voltage (V)"].values,
                'power': df["Power (W)"].values
            }
            
            self.loaded_curves.append(curve_data)
            display_name = f"{mode} {sense} - {filename}"
            self.file_listbox.insert(tk.END, display_name)
            self.update_plot()
            self.update_statistics()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV file:\n{file_path}\n\nError: {e}")
    
    def remove_selected_file(self):
        """Remove the selected file from the comparison."""
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.loaded_curves.pop(index)
            self.file_listbox.delete(index)
            self.update_plot()
            self.update_statistics()
    
    def clear_all_files(self):
        """Clear all loaded files."""
        self.loaded_curves.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_plot()
        self.update_statistics()
    
    def update_plot(self):
        """Update the comparison plot with all loaded curves."""
        # Clear previous plots
        self.ax1.clear()
        self.ax2.clear()
        
        if not self.loaded_curves:
            self.ax1.text(0.5, 0.5, "No data loaded\nClick 'Browse Recent Measurements' to load your existing CSV files", 
                         ha='center', va='center', transform=self.ax1.transAxes)
            self.ax2.text(0.5, 0.5, "No data loaded", ha='center', va='center', transform=self.ax2.transAxes)
            self.canvas.draw()
            return
        
        # Color palette for different curves - more distinct colors
        colors = ['blue', 'red', 'green', 'purple', 'orange', 'brown', 'pink', 'gray', 'olive', 'cyan']
        # Larger and more distinct markers
        markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p', '*', 'h', '+', 'x']
        # Different line styles for extra distinction
        linestyles = ['-', '--', '-.', ':', '-', '--', '-.', ':', '-', '--']
        
        # Plot I-V curves if enabled
        if self.show_iv_var.get():
            self.ax1.set_xlabel("Voltage (V)")
            self.ax1.set_ylabel("Current (A)")
            self.ax1.set_title("I-V Curve Comparison")
            self.ax1.grid(True, alpha=0.3)
            
            for i, curve in enumerate(self.loaded_curves):
                color = colors[i % len(colors)]
                marker = markers[i % len(markers)]
                linestyle = linestyles[i % len(linestyles)]
                label = f"{curve['mode']} {curve['sense']}"
                
                # Use absolute values for current to match the main application display
                current_abs = [abs(c) for c in curve['current']]
                
                self.ax1.plot(curve['voltage'], current_abs, 
                             color=color, marker=marker, markersize=8, linewidth=3,
                             linestyle=linestyle, label=label, alpha=0.8, 
                             markeredgewidth=1, markeredgecolor='black')
            
            # Set both X and Y axes to start from 0 (like in main application)
            self.ax1.set_xlim(left=0)
            self.ax1.set_ylim(bottom=0)
            self.ax1.legend()
        
        # Plot P-V curves if enabled
        if self.show_pv_var.get():
            self.ax2.set_xlabel("Voltage (V)")
            self.ax2.set_ylabel("Power (W)")
            self.ax2.set_title("P-V Curve Comparison")
            self.ax2.grid(True, alpha=0.3)
            
            for i, curve in enumerate(self.loaded_curves):
                color = colors[i % len(colors)]
                marker = markers[i % len(markers)]
                linestyle = linestyles[i % len(linestyles)]
                label = f"{curve['mode']} {curve['sense']}"
                
                # Use absolute values for power calculation to ensure positive power
                power_abs = [abs(p) for p in curve['power']]
                
                self.ax2.plot(curve['voltage'], power_abs, 
                             color=color, marker=marker, markersize=8, linewidth=3,
                             linestyle=linestyle, label=label, alpha=0.8,
                             markeredgewidth=1, markeredgecolor='black')
                
                # Mark maximum power point with larger star
                max_power_idx = power_abs.index(max(power_abs))
                max_power = power_abs[max_power_idx]
                max_power_voltage = curve['voltage'][max_power_idx]
                
                self.ax2.plot(max_power_voltage, max_power, 
                             color=color, marker='*', markersize=15, 
                             markeredgecolor='black', markeredgewidth=2)
            
            # Set both X and Y axes to start from 0 (like in main application)
            self.ax2.set_xlim(left=0)
            self.ax2.set_ylim(bottom=0)
            self.ax2.legend()
        
        self.figure.tight_layout()
        self.canvas.draw()
    
    def update_statistics(self):
        """Update the statistics display."""
        self.stats_text.delete(1.0, tk.END)
        
        if not self.loaded_curves:
            self.stats_text.insert(tk.END, "No data loaded")
            return
        
        stats_text = "Curve Statistics:\n" + "="*30 + "\n\n"
        
        for i, curve in enumerate(self.loaded_curves):
            stats_text += f"Curve {i+1}: {curve['mode']} {curve['sense']}\n"
            stats_text += f"File: {curve['filename']}\n"
            
            try:
                # Convert to numpy arrays for easier manipulation
                current_array = curve['current']
                voltage_array = curve['voltage']
                power_array = curve['power']
                
                # Use absolute values to match main application display
                current_abs = [abs(float(c)) for c in current_array]
                power_abs = [abs(float(p)) for p in power_array]
                voltage_vals = [float(v) for v in voltage_array]
                
                # Calculate key parameters with proper error handling
                max_power_idx = power_abs.index(max(power_abs))
                pmp = power_abs[max_power_idx]
                vmp = voltage_vals[max_power_idx]
                imp = current_abs[max_power_idx]
                
                # Find Voc and Isc more robustly
                min_current_idx = current_abs.index(min(current_abs))
                voc = voltage_vals[min_current_idx]
                
                min_voltage_idx = voltage_vals.index(min(voltage_vals))
                isc = current_abs[min_voltage_idx]
                
                # Calculate fill factor with proper type conversion
                if (voc * isc) > 0:
                    fill_factor = (pmp / (voc * isc)) * 100
                else:
                    fill_factor = 0
                
                stats_text += f"Pmp: {pmp:.3f} W\n"
                stats_text += f"Vmp: {vmp:.3f} V\n"
                stats_text += f"Imp: {imp:.3f} A\n"
                stats_text += f"Voc: {voc:.3f} V\n"
                stats_text += f"Isc: {isc:.3f} A\n"
                stats_text += f"FF: {fill_factor:.1f}%\n"
                
            except (ValueError, TypeError, IndexError) as e:
                stats_text += f"Error calculating parameters: {e}\n"
            
            stats_text += "-"*25 + "\n\n"
        
        self.stats_text.insert(tk.END, stats_text)
    
    def export_comparison(self):
        """Export the comparison plot and statistics."""
        if not self.loaded_curves:
            messagebox.showwarning("Warning", "No data to export")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Export plot
        plot_path = filedialog.asksaveasfilename(
            title="Save Comparison Plot",
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf")],
            initialfile=f"IV_Comparison_{timestamp}.png"
        )
        
        if plot_path:
            try:
                self.figure.savefig(plot_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Export", f"Plot saved to:\n{plot_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save plot:\n{e}")

if __name__ == "__main__":
    # Start the main application window
    root = tk.Tk()
    root.geometry("950x850")
    root.resizable(True, True)
    app = IVAppCC(root)
    root.mainloop()