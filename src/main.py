# Import required libraries for GUI, plotting, instrument control, and data handling
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
import pandas as pd  # For advanced CSV data handling and analysis

class IVAppCC:
    """
    Main application class for I-V curve measurement using electronic loads.
    Supports both Constant Current (CC) and Constant Voltage (CV) modes with 2-Wire or 4-Wire sensing.
    Features real-time plotting, data saving, and comprehensive safety protections.
    """
    
    def __init__(self, root):
        """
        Initialize the main application window and all GUI components.
        
        Args:
            root: The main tkinter window object
        """
        self.root = root
        self.root.title("I-V Curve Measurement (CC/CV Mode)")

        # VISA instrument selection dropdown - detects available instruments
        tk.Label(root, text="Select Instrument:").grid(row=0, column=0, sticky="e")
        try:
            # Initialize VISA resource manager to communicate with instruments
            self.rm = pyvisa.ResourceManager()
            real_instr = list(self.rm.list_resources())
        except Exception:
            # Handle case where no VISA drivers are installed or no instruments connected
            self.rm = None
            real_instr = []
        
        # Create dropdown with simulated instrument option plus any real instruments found
        self.instr_list = ["Simulated Instrument"] + real_instr
        self.instr_var = tk.StringVar()
        self.instr_dropdown = ttk.Combobox(root, textvariable=self.instr_var, values=self.instr_list, state="readonly")
        self.instr_dropdown.grid(row=0, column=1, columnspan=2, sticky="ew")
        self.instr_dropdown.current(0)  # Default to simulated instrument

        # Sense mode selection (2-Wire or 4-Wire) - affects measurement accuracy
        tk.Label(root, text="Sense Mode:").grid(row=1, column=0, sticky="e")
        self.sense_mode_var = tk.StringVar()
        self.sense_mode_dropdown = ttk.Combobox(root, textvariable=self.sense_mode_var, values=["2-Wire", "4-Wire"], state="readonly")
        self.sense_mode_dropdown.grid(row=1, column=1, columnspan=2, sticky="ew")
        self.sense_mode_dropdown.current(0)  # Default to 2-Wire

        # Operation mode selection (CC or CV) - determines sweep variable
        tk.Label(root, text="Operation Mode:").grid(row=2, column=0, sticky="e")
        self.mode_var = tk.StringVar()
        self.mode_dropdown = ttk.Combobox(root, textvariable=self.mode_var, values=["CC", "CV"], state="readonly")
        self.mode_dropdown.grid(row=2, column=1, columnspan=2, sticky="ew")
        self.mode_dropdown.current(0)  # Default to CC mode
        # Add callback to update labels when mode changes
        self.mode_var.trace("w", self.update_labels)

        # Input fields for sweep parameters - labels change based on selected mode
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

        # Protection settings - critical for solar cell safety
        tk.Label(root, text="Voltage Limit (V):").grid(row=6, column=0, sticky="e")
        self.voltage_limit_entry = tk.Entry(root)
        self.voltage_limit_entry.grid(row=6, column=1, columnspan=2, sticky="ew")

        tk.Label(root, text="Current Limit (A):").grid(row=7, column=0, sticky="e")
        self.current_limit_entry = tk.Entry(root)
        self.current_limit_entry.grid(row=7, column=1, columnspan=2, sticky="ew")

        # Sleep time between measurement steps - allows settling time
        tk.Label(root, text="Step Delay (s):").grid(row=8, column=0, sticky="e")
        self.sleep_time_entry = tk.Entry(root)
        self.sleep_time_entry.grid(row=8, column=1, columnspan=2, sticky="ew")
        self.sleep_time_entry.insert(0, "0.5")  # Default 500ms delay

        # Save options for data and plots
        self.save_csv_var = tk.BooleanVar()
        self.save_csv_check = tk.Checkbutton(root, text="Save CSV data", variable=self.save_csv_var)
        self.save_csv_check.grid(row=9, column=0, columnspan=3, sticky="w", padx=5)

        self.save_png_var = tk.BooleanVar()
        self.save_png_check = tk.Checkbutton(root, text="Save plot as PNG", variable=self.save_png_var)
        self.save_png_check.grid(row=10, column=0, columnspan=3, sticky="w", padx=5)

        # Output directory configuration - organized by date in project root
        self.output_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'output'))
        os.makedirs(self.output_dir, exist_ok=True)
        self.folder_label = tk.Label(root, text=f"Output directory: {self.output_dir}")
        self.folder_label.grid(row=14, column=0, columnspan=2, sticky="w", padx=5, pady=(5, 0))
        self.folder_button = tk.Button(root, text="Change...", command=self.choose_output_dir)
        self.folder_button.grid(row=14, column=2, sticky="e", padx=5, pady=(5, 0))

        # Control buttons for sweep operation
        self.start_button = tk.Button(root, text="Start Sweep", command=self.start_sweep_thread)
        self.start_button.grid(row=11, column=0, columnspan=2, pady=10)

        self.stop_button = tk.Button(root, text="Stop", command=self.request_stop, state="normal")
        self.stop_button.grid(row=11, column=2, pady=10)

        # Comparison button to open curve comparison window
        self.compare_button = tk.Button(root, text="Compare Curves", command=self.open_comparison_window)
        self.compare_button.grid(row=11, column=3, pady=10)

        # Progress bar to show sweep completion status
        self.progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')
        self.progress.grid(row=12, column=0, columnspan=3, sticky="ew", padx=5)

        # Matplotlib plot area for real-time I-V and P-V curve display
        self.figure = plt.Figure(figsize=(7, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=13, column=0, columnspan=3, sticky="nsew")

        # Configure GUI resizing behavior - plot area expands with window
        root.grid_rowconfigure(13, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # State flags for sweep control and thread safety
        self.stop_requested = False
        self.sweep_running = False

        # Keyboard shortcuts for convenience
        self.root.bind('<Return>', self.on_enter)    # Enter to start sweep
        self.root.bind('<Escape>', self.on_escape)   # Escape to stop sweep

        # Load previously saved settings on startup
        self.load_settings()

    def choose_output_dir(self):
        """
        Open a file dialog to allow user to select a different output directory.
        Updates the folder label with the new selection.
        """
        new_dir = filedialog.askdirectory(initialdir=os.getcwd(), title="Choose output directory")
        if new_dir:
            self.output_dir = new_dir
            self.folder_label.config(text=f"Output directory: {self.output_dir}")

    def on_enter(self, event):
        """
        Keyboard shortcut handler for Enter key.
        Starts sweep if not already running.
        
        Args:
            event: Tkinter key event object
        """
        if not self.sweep_running:
            self.start_sweep_thread()

    def on_escape(self, event):
        """
        Keyboard shortcut handler for Escape key.
        Stops sweep if currently running.
        
        Args:
            event: Tkinter key event object
        """
        if self.sweep_running:
            self.request_stop()

    def start_sweep_thread(self):
        """
        Start the measurement sweep in a separate thread to prevent GUI freezing.
        Uses daemon thread to ensure clean program termination.
        """
        if self.sweep_running:
            return  # Prevent multiple simultaneous sweeps
        
        self.sweep_running = True
        thread = threading.Thread(target=self.start_sweep)
        thread.daemon = True  # Thread will terminate when main program exits
        thread.start()

    def request_stop(self):
        """
        Set the stop flag to request sweep interruption.
        The sweep loop checks this flag and stops gracefully.
        """
        self.stop_requested = True

    def update_labels(self, *args):
        """
        Update the input field labels based on the selected operation mode.
        CC mode uses current units, CV mode uses voltage units.
        
        Args:
            *args: Variable arguments from tkinter trace callback (unused)
        """
        selected_mode = self.mode_var.get()
        if selected_mode == "CC":
            # Constant Current mode - sweeping current, measuring voltage
            self.start_label.config(text="Start Current (A):")
            self.end_label.config(text="End Current (A):")
            self.step_label.config(text="Step Current (A):")
        else:
            # Constant Voltage mode - sweeping voltage, measuring current
            self.start_label.config(text="Start Voltage (V):")
            self.end_label.config(text="End Voltage (V):")
            self.step_label.config(text="Step Voltage (V):")

    def save_settings(self):
        """
        Save current GUI settings to a JSON file for persistence between sessions.
        Includes all user inputs, mode selections, and preferences.
        """
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
        """
        Load previously saved GUI settings from JSON file.
        Restores all user inputs and preferences from last session.
        """
        if os.path.exists("last_settings.json"):
            with open("last_settings.json", "r") as f:
                settings = json.load(f)
            
            # Restore all input field values
            self.start_current_entry.delete(0, tk.END)
            self.start_current_entry.insert(0, settings.get("start", ""))
            self.end_current_entry.delete(0, tk.END)
            self.end_current_entry.insert(0, settings.get("end", ""))
            self.step_current_entry.delete(0, tk.END)
            self.step_current_entry.insert(0, settings.get("step", ""))
            self.voltage_limit_entry.delete(0, tk.END)
            self.voltage_limit_entry.insert(0, settings.get("voltage_limit", ""))
            self.current_limit_entry.delete(0, tk.END)
            self.current_limit_entry.insert(0, settings.get("current_limit", ""))
            self.sleep_time_entry.delete(0, tk.END)
            self.sleep_time_entry.insert(0, settings.get("sleep_time", ""))
            
            # Restore dropdown selections
            self.mode_var.set(settings.get("mode", "CC"))
            self.sense_mode_var.set(settings.get("sense", "2-Wire"))
            if settings.get("instr") in self.instr_list:
                self.instr_var.set(settings.get("instr"))
            
            # Restore checkboxes and output directory
            self.save_csv_var.set(settings.get("save_csv", False))
            self.save_png_var.set(settings.get("save_png", False))
            if os.path.isdir(settings.get("output_dir", "")):
                self.output_dir = settings["output_dir"]
                self.folder_label.config(text=f"Output directory: {self.output_dir}")

    def start_sweep(self):
        """
        Main function to perform the I-V sweep measurement.
        Handles instrument communication, real-time plotting, data collection,
        safety monitoring, and file saving. Runs in separate thread.
        """
        self.stop_requested = False  # Reset stop flag for new sweep
        
        try:
            # Parse and validate user input values from GUI
            i_start = float(self.start_current_entry.get())
            i_end = float(self.end_current_entry.get())
            i_step = float(self.step_current_entry.get())
            
            # Parse protection limits (optional but recommended)
            voltage_limit = self.voltage_limit_entry.get()
            current_limit = self.current_limit_entry.get()
            voltage_limit = float(voltage_limit) if voltage_limit else None
            current_limit = float(current_limit) if current_limit else None
            
            sleep_time = float(self.sleep_time_entry.get())
            
        except ValueError:
            # Handle invalid numeric input
            messagebox.showerror("Input Error", "Please enter valid numbers.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return

        selected_mode = self.mode_var.get()
        
        # Critical safety checks - protection limits are mandatory for solar cell safety
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

        # Validate step size
        if i_step == 0:
            messagebox.showerror("Input Error", "Step value cannot be zero.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return

        # Validate instrument selection
        if not self.instr_var.get():
            messagebox.showerror("Connection Error", "No instrument selected.")
            self.sweep_running = False
            self.start_button.config(state='normal')
            return

        # Disable start button during sweep to prevent conflicts
        self.start_button.config(state='disabled')
        instrument_address = self.instr_var.get()

        try:
            # Initialize instrument connection (real hardware or simulation)
            if instrument_address == "Simulated Instrument":
                load = self.create_simulated_instrument()
            else:
                # Connect to real instrument via VISA
                load = self.rm.open_resource(instrument_address)
                load.timeout = 5000  # 5 second timeout for commands
                load.write("*RST")   # Reset instrument to known state
                load.write("*CLS")   # Clear status registers

            # Configure instrument operating mode
            selected_mode = self.mode_var.get()
            mode_mapping = {"CC": "CURR", "CV": "VOLT"}
            if load.query("FUNC?").strip() != mode_mapping[selected_mode]:
                load.write(f"FUNC {mode_mapping[selected_mode]}")

            # Configure safety protection limits
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

            # Configure sensing mode (affects measurement accuracy)
            sense_command = "REM:SENS ON" if self.sense_mode_var.get() == "4-Wire" else "REM:SENS OFF"
            load.write(sense_command)
            time.sleep(0.5)  # Allow settings to take effect

            # Enable instrument input after all configuration is complete
            load.write("INPUT ON")

            # Prepare dual-axis plot (I-V curve on left axis, P-V curve on right axis)
            self.ax.clear()
            if hasattr(self, 'ax2'):
                self.figure.delaxes(self.ax2)
            self.ax2 = self.ax.twinx()  # Create secondary y-axis for power

            # Remove any existing plot elements from previous sweeps
            for attr in ['line_iv', 'line_power', 'pmp_annotation', 'pmp_point', 'vmp_annotation', 'vmp_point', 'summary_annotation']:
                if hasattr(self, attr):
                    try:
                        getattr(self, attr).remove()
                    except Exception:
                        pass
                    delattr(self, attr)         

            # Configure plot axes and appearance
            self.ax.set_xlabel("Voltage (V)")
            self.ax.set_ylabel("Current (A)", color='b')
            self.ax.tick_params(axis='y', labelcolor='b')
            self.ax.grid(True)
            self.ax2.yaxis.set_label_position("right")
            self.ax2.yaxis.tick_right()
            self.ax2.set_ylabel("Power (W)", color='r')
            self.ax2.tick_params(axis='y', labelcolor='r')
            self.canvas.draw()

            # Initialize data storage lists
            currents, voltages, powers = [], [], []

            # Configure sweep parameters based on operating mode
            if selected_mode == "CC":
                # Constant Current mode: sweep current, measure voltage
                sweep_start = i_start
                sweep_end = i_end
                sweep_step = i_step if i_end >= i_start else -i_step
                setpoint_cmd = lambda v: load.write(f"CURR {v:.3f}")
            else:
                # Constant Voltage mode: sweep voltage, measure current
                sweep_start = i_start
                sweep_end = i_end
                sweep_step = i_step if i_end >= i_start else -i_step
                setpoint_cmd = lambda v: load.write(f"VOLT {v:.3f}")

            # Calculate sweep parameters
            value = sweep_start
            total_steps = int(abs((sweep_end - sweep_start) / sweep_step)) + 1
            self.progress["maximum"] = total_steps
            self.progress["value"] = 0
            print(f"total_steps = {total_steps}, sweep_start = {sweep_start}, sweep_end = {sweep_end}, sweep_step = {sweep_step}")

            # Set initial setpoint and allow settling
            setpoint_cmd(sweep_start)
            time.sleep(sleep_time)

            # Ensure input is enabled before starting measurements
            load.write("INPUT ON")
            time.sleep(0.2)

            # Main measurement loop
            for count in range(total_steps):
                # Check for user-requested stop
                if self.stop_requested:
                    messagebox.showinfo("Sweep Stopped", "Sweep was stopped by the user.")
                    break
                
                try:
                    # Set new setpoint and allow settling
                    setpoint_cmd(value)
                    time.sleep(sleep_time)
                    
                    # Read measurements from instrument
                    voltage = float(load.query("MEAS:VOLT?"))
                    actual_current = float(load.query("MEAS:CURR?"))
                    power = voltage * actual_current

                    # Safety protection checks - stop if limits exceeded
                    if voltage_limit is not None and voltage > voltage_limit:
                        raise Exception("Voltage exceeded protection limit.")
                    if current_limit is not None and actual_current > current_limit:
                        raise Exception("Current exceeded protection limit.")

                    # Debug output for monitoring
                    print(f"Protection check: V={voltage} (limit {voltage_limit}), I={actual_current} (limit {current_limit})")
                    print(f"Setpoint: {value:.3f} V, Measured: {voltage:.3f} V, {actual_current:.3f} A")

                    # Store data point (avoid duplicates within tolerance)
                    EPS = 1e-4
                    if len(currents) == 0 or abs(actual_current - currents[-1]) > EPS or abs(voltage - voltages[-1]) > EPS:
                        currents.append(actual_current)
                        voltages.append(voltage)
                        powers.append(power)

                    # Update I-V curve plot in real-time
                    if hasattr(self, 'line_iv'):
                        self.line_iv.remove()
                        del self.line_iv
                    self.line_iv, = self.ax.plot(voltages, currents, label="I-V Curve", color='blue')

                    # Update P-V curve plot in real-time
                    if hasattr(self, 'line_power'):
                        self.line_power.remove()
                        del self.line_power
                    self.line_power, = self.ax2.plot(voltages, powers, label="P-V Curve", color='red')

                    # Auto-scale axes to fit data
                    self.ax.relim()
                    self.ax.autoscale_view()
                    self.ax2.relim()
                    self.ax2.autoscale_view()

                    # Set X axis to always start at 0V for consistency
                    if voltages:
                        v_max = max(voltages)
                        self.ax.set_xlim(left=0, right=v_max * 1.0105)
                    else:
                        self.ax.set_xlim(left=0)

                    # Update display and progress
                    self.canvas.draw()
                    self.root.update_idletasks()
                    self.progress["value"] = count + 1
                    
                except Exception as e:
                    # Handle measurement errors or protection trips
                    print(f"Exception in sweep loop: {e}")
                    messagebox.showwarning("Protection Triggered", f"Sweep stopped: {e}")
                    break
                
                # Advance to next setpoint
                value += sweep_step

            # Clean shutdown - turn off load and disconnect
            load.write("INPUT OFF")
            load.close()

            # Final plot update with complete data
            if voltages and currents and hasattr(self, 'line_iv'):
                self.line_iv.set_data(voltages, currents)
            if voltages and powers and hasattr(self, 'line_power'):
                self.line_power.set_data(voltages, powers)
            
            # Finalize plot appearance
            self.ax.set_xlabel("Voltage (V)")
            self.ax.set_ylabel("Current (A)", color='b')
            self.ax2.set_ylabel("Power (W)", color='r')
            self.ax.relim()
            self.ax.autoscale_view()
            self.ax2.relim()
            self.ax2.autoscale_view()
            self.canvas.draw()

            # Calculate and display key photovoltaic parameters
            if powers:
                pmp = max(powers)           # Maximum power point
                idx = powers.index(pmp)
                vmp = voltages[idx]         # Voltage at maximum power
                imp = currents[idx]         # Current at maximum power
                summary_text = f"Pmp = {pmp:.2f} W   Vmp = {vmp:.2f} V   Imp = {imp:.2f} A"
            else:
                pmp = vmp = imp = None
                summary_text = "Sweep completed with no power data."

            # Remove previous summary annotation
            if hasattr(self, 'summary_annotation'):
                try:
                    self.summary_annotation.remove()
                except Exception:
                    pass
                del self.summary_annotation

            # Add summary text box above the plot
            self.summary_annotation = self.ax.annotate(
                summary_text,
                xy=(0.5, 1.08), xycoords='axes fraction',
                ha='center', va='bottom',
                fontsize=14, color='purple',
                bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="purple", lw=2)
            )
            self.canvas.draw()

            # Prepare file naming and metadata
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            selected_mode = self.mode_var.get()
            sense_mode = self.sense_mode_var.get()
            base_filename = f"IV_Sweep_{selected_mode}_{sense_mode}_{timestamp}"

            # Create parameter list for CSV metadata
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

            # Create date-organized output directory
            today_str = datetime.now().strftime("%Y-%m-%d")
            day_output_dir = os.path.join(self.output_dir, today_str)
            os.makedirs(day_output_dir, exist_ok=True)

            # Save CSV data file if requested
            if self.save_csv_var.get():
                csv_path = os.path.join(day_output_dir, f"{base_filename}.csv")
                with open(csv_path, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    # Write measurement data
                    writer.writerow(["Current (A)", "Voltage (V)", "Power (W)"])
                    for i in range(len(currents)):
                        writer.writerow([currents[i], voltages[i], powers[i]])
                    # Write metadata section
                    writer.writerow([])
                    writer.writerow(["Parameter", "Value"])
                    for param, value in params:
                        writer.writerow([param, value])
                print(f"Data saved to {csv_path}")

            # Highlight maximum power point on the plot
            if powers and voltages and currents:
                pmp = max(powers)
                idx = powers.index(pmp)
                vmp = voltages[idx]
                imp = currents[idx]
                
                # Add prominent marker at Pmp on P-V curve
                self.ax2.plot(vmp, pmp, 'ro', markersize=12, label="Pmp")
                
                # Add annotation with arrow pointing to Pmp
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
                self.canvas.draw()

            # Save plot as PNG if requested
            if self.save_png_var.get():
                png_path = os.path.join(day_output_dir, f"{base_filename}.png")
                self.figure.savefig(png_path)
                print(f"Plot saved to {png_path}")

            # Display completion message with key results
            message = f"Sweep completed.\nPmp = {pmp:.2f} W\nVmp = {vmp:.2f} V\nImp = {imp:.2f} A" if pmp else "Sweep completed with no power data."
            messagebox.showinfo("Sweep Complete", message)

        except Exception as e:
            # Handle any unexpected errors during sweep
            messagebox.showerror("Error", f"An error occurred:\n{e}")

        finally:
            # Always execute cleanup regardless of success or failure
            self.save_settings()           # Preserve user settings
            self.sweep_running = False     # Reset sweep state
            self.start_button.config(state='normal')  # Re-enable start button

    def create_simulated_instrument(self):
        """
        Create a simulated electronic load for testing without physical hardware.
        Implements realistic solar cell I-V characteristics using diode equation.
        
        Returns:
            SimulatedInstrument: Object that mimics real instrument behavior
        """
        class SimulatedInstrument:
            """
            Simulated electronic load that responds to SCPI commands.
            Models a solar cell with realistic I-V characteristics.
            """
            
            def __init__(self):
                """Initialize instrument state with default parameters."""
                self.state = {
                    "FUNC": "CURR",           # Operating mode (CURR or VOLT)
                    "current": 0.0,           # Set current value
                    "voltage": 0.0,           # Set voltage value
                    "VOLT_PROT_ON": False,    # Voltage protection enable
                    "VOLT_PROT": None,        # Voltage protection limit
                    "CURR_PROT_ON": False,    # Current protection enable
                    "CURR_PROT": None,        # Current protection limit
                }

            def write(self, command):
                """
                Process SCPI commands sent to the instrument.
                
                Args:
                    command (str): SCPI command string
                """
                # Parse function selection commands
                if "FUNC" in command:
                    self.state["FUNC"] = command.split()[-1]
                # Parse setpoint commands
                elif "CURR" in command and "CURR:PROT" not in command:
                    self.state["current"] = float(command.split()[-1])
                elif "VOLT" in command and "VOLT:PROT" not in command:
                    self.state["voltage"] = float(command.split()[-1])
                # Parse protection commands
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
                """
                Process SCPI queries and return simulated measurements.
                Uses solar cell model: I = Isc * (1 - exp((V - Voc)/(n*Vt)))
                
                Args:
                    command (str): SCPI query string
                    
                Returns:
                    str: Simulated measurement result
                """
                # Solar cell model parameters
                Isc = 5.0     # Short circuit current (A)
                Voc = 25      # Open circuit voltage (V) - high for demonstration
                n = 1.5       # Ideality factor
                Vt = 0.7      # Thermal voltage (V)

                if "MEAS:VOLT?" in command:
                    # Voltage measurement query
                    if self.state["FUNC"] == "CURR":
                        # CC mode: calculate voltage for given current
                        I = self.state["current"]
                        # Numerically invert diode equation: V = Voc + n*Vt*ln(1 - I/Isc)
                        V = Voc + n * Vt * math.log(1 - I / Isc) if I < Isc else 0
                        # Check voltage protection
                        if self.state["VOLT_PROT_ON"] and self.state["VOLT_PROT"] is not None and V > self.state["VOLT_PROT"]:
                            return str(self.state["VOLT_PROT"] + 5)  # Simulate protection trip
                        return str(max(V, 0))
                    return str(self.state["voltage"])
                    
                elif "MEAS:CURR?" in command:
                    # Current measurement query
                    if self.state["FUNC"] == "VOLT":
                        # CV mode: calculate current for given voltage using diode equation
                        V = self.state["voltage"]
                        I = Isc * (1 - math.exp((V - Voc) / (n * Vt)))
                        I = max(I, 0)  # Ensure non-negative current
                        # Check current protection
                        if self.state["CURR_PROT_ON"] and self.state["CURR_PROT"] is not None and I > self.state["CURR_PROT"]:
                            return str(self.state["CURR_PROT"] + 5)  # Simulate protection trip
                        return str(I)
                    return str(self.state["current"])
                    
                elif "FUNC?" in command:
                    # Function query
                    return self.state.get("FUNC", "CURR")
                elif "STAT:QUES:COND?" in command:
                    # Status query (always return no errors for simulation)
                    return "0"
                return "0"

            def close(self):
                """Close instrument connection (no-op for simulation)."""
                pass

        return SimulatedInstrument()

    def open_comparison_window(self):
        """
        Open a new window for comparing multiple I-V curves.
        Creates a separate application for loading and analyzing historical data.
        """
        comparison_window = tk.Toplevel(self.root)
        comparison_window.title("I-V Curve Comparison")
        comparison_window.geometry("1200x800")
        
        # Create comparison application instance
        ComparisonApp(comparison_window, self.output_dir)

class ComparisonApp:
    """
    Secondary application window for comparing multiple I-V curves.
    Allows loading historical CSV files and displaying them together for analysis.
    """
    
    def __init__(self, root, default_output_dir):
        """
        Initialize the comparison window with file browser and plotting capabilities.
        
        Args:
            root: Tkinter window object for the comparison window
            default_output_dir: Default directory to browse for CSV files
        """
        self.root = root
        self.output_dir = default_output_dir
        self.loaded_curves = []  # Storage for loaded curve data dictionaries
        
        # Create main layout with control panel and plot area
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left side control panel
        control_frame = tk.Frame(main_frame)
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # File selection section
        tk.Label(control_frame, text="Load I-V Data Files:", font=("Arial", 12, "bold")).pack(anchor="w")
        
        # Quick access button for recent measurements
        tk.Button(control_frame, text="Browse Recent Measurements", 
                 command=self.browse_recent_measurements, bg="lightgreen").pack(pady=5)
        
        # List box to display currently loaded files
        self.file_listbox = tk.Listbox(control_frame, height=8, width=40)
        self.file_listbox.pack(pady=(5, 0))
        
        # File management buttons
        button_frame = tk.Frame(control_frame)
        button_frame.pack(pady=5)
        
        tk.Button(button_frame, text="Add CSV File", command=self.add_csv_file).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(button_frame, text="Remove Selected", command=self.remove_selected_file).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(button_frame, text="Clear All", command=self.clear_all_files).pack(side=tk.LEFT)
        
        # Plot configuration options
        tk.Label(control_frame, text="Plot Options:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(20, 5))
        
        # Checkboxes to control which curves are displayed
        self.show_iv_var = tk.BooleanVar(value=True)
        self.show_pv_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(control_frame, text="Show I-V Curves", variable=self.show_iv_var, command=self.update_plot).pack(anchor="w")
        tk.Checkbutton(control_frame, text="Show P-V Curves", variable=self.show_pv_var, command=self.update_plot).pack(anchor="w")
        
        # Plot update and export controls
        tk.Button(control_frame, text="Update Plot", command=self.update_plot, 
                 bg="lightblue", font=("Arial", 10, "bold")).pack(pady=10)
        
        tk.Button(control_frame, text="Export Comparison", command=self.export_comparison).pack(pady=(0, 10))
        
        # Statistics display area
        tk.Label(control_frame, text="Statistics:", font=("Arial", 12, "bold")).pack(anchor="w", pady=(20, 5))
        self.stats_text = tk.Text(control_frame, height=10, width=40, font=("Courier", 9))
        self.stats_text.pack()
        
        # Right side plot area
        plot_frame = tk.Frame(main_frame)
        plot_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Create matplotlib figure with two subplots
        self.figure = plt.Figure(figsize=(10, 8), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Initialize subplot areas
        self.ax1 = self.figure.add_subplot(211)  # Top subplot for I-V curves
        self.ax2 = self.figure.add_subplot(212)  # Bottom subplot for P-V curves
        self.figure.tight_layout()
        self.canvas.draw()
    
    def browse_recent_measurements(self):
        """
        Open a file browser window showing recent measurements organized by date.
        Allows multiple file selection for batch loading.
        """
        # Verify output directory exists
        if not os.path.exists(self.output_dir):
            messagebox.showwarning("Warning", f"Output directory not found: {self.output_dir}")
            return
        
        # Create file selection dialog
        selection_window = tk.Toplevel(self.root)
        selection_window.title("Select Recent Measurements")
        selection_window.geometry("600x400")
        
        # Recursively find all CSV files in output directory
        csv_files = []
        for root, dirs, files in os.walk(self.output_dir):
            for file in files:
                if file.endswith('.csv'):
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, self.output_dir)
                    csv_files.append((rel_path, full_path))
        
        # Sort files by modification time (newest first)
        csv_files.sort(key=lambda x: os.path.getmtime(x[1]), reverse=True)
        
        # Create file list interface
        tk.Label(selection_window, text="Recent Measurements (newest first):", 
                font=("Arial", 12, "bold")).pack(pady=5)
        
        # Listbox with scrollbar for file selection
        list_frame = tk.Frame(selection_window)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        file_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, selectmode=tk.MULTIPLE)
        file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=file_listbox.yview)
        
        # Populate list with found CSV files
        for rel_path, full_path in csv_files:
            file_listbox.insert(tk.END, rel_path)
        
        # Control buttons for file selection
        button_frame = tk.Frame(selection_window)
        button_frame.pack(pady=10)
        
        def load_selected():
            """Load all selected files and close selection window."""
            selected_indices = file_listbox.curselection()
            if not selected_indices:
                messagebox.showwarning("Warning", "Please select at least one file.")
                return
            
            # Load each selected file
            for idx in selected_indices:
                _, full_path = csv_files[idx]
                self.load_csv_file(full_path)
            
            selection_window.destroy()
        
        tk.Button(button_frame, text="Load Selected", command=load_selected, 
                 bg="lightgreen").pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=selection_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def add_csv_file(self):
        """
        Open file dialog to manually select a single CSV file for comparison.
        """
        file_path = filedialog.askopenfilename(
            title="Select I-V Data CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=self.output_dir
        )
        
        if file_path:
            self.load_csv_file(file_path)
    
    def load_csv_file(self, file_path):
        """
        Load and parse a CSV file containing I-V measurement data.
        Performs data validation, cleaning, and metadata extraction.
        
        Args:
            file_path (str): Full path to the CSV file to load
        """
        try:
            # Read CSV file using pandas for robust parsing
            df = pd.read_csv(file_path)
            
            # Validate required columns are present
            required_cols = ["Current (A)", "Voltage (V)", "Power (W)"]
            if not all(col in df.columns for col in required_cols):
                messagebox.showerror("Error", f"CSV file must contain columns: {', '.join(required_cols)}")
                return
            
            # Data cleaning and validation process
            df = df.dropna()  # Remove rows with missing values
            
            # Convert all measurement columns to numeric, handling text/headers
            df["Current (A)"] = pd.to_numeric(df["Current (A)"], errors='coerce')
            df["Voltage (V)"] = pd.to_numeric(df["Voltage (V)"], errors='coerce')
            df["Power (W)"] = pd.to_numeric(df["Power (W)"], errors='coerce')
            
            # Remove rows that couldn't be converted to numbers
            df = df.dropna()
            
            # Remove zero-only rows (often header repetitions or noise)
            df = df[(df["Current (A)"] != 0) | (df["Voltage (V)"] != 0) | (df["Power (W)"] != 0)]
            
            # Validate that data remains after cleaning
            if df.empty:
                messagebox.showerror("Error", "No valid numeric data found in CSV file")
                return
            
            # Extract metadata from filename using naming convention
            filename = os.path.basename(file_path)
            
            # Parse operating mode from filename
            mode = "Unknown"
            sense = "Unknown"
            
            if "_CC_" in filename:
                mode = "CC"  # Constant Current mode
            elif "_CV_" in filename:
                mode = "CV"  # Constant Voltage mode
            
            # Parse sensing mode from filename
            if "_4-Wire_" in filename:
                sense = "4-Wire"  # Kelvin sensing
            elif "_2-Wire_" in filename:
                sense = "2-Wire"  # Standard sensing
            
            # Fallback mode detection using data characteristics
            if mode == "Unknown":
                # Heuristic: larger voltage range suggests CV mode
                voltage_range = df["Voltage (V)"].max() - df["Voltage (V)"].min()
                current_range = df["Current (A)"].max() - df["Current (A)"].min()
                mode = "CV" if voltage_range > current_range else "CC"
            
            # Create curve data structure
            curve_data = {
                'file_path': file_path,
                'filename': filename,
                'mode': mode,
                'sense': sense,
                'current': df["Current (A)"].values,
                'voltage': df["Voltage (V)"].values,
                'power': df["Power (W)"].values
            }
            
            # Add to loaded curves and update displays
            self.loaded_curves.append(curve_data)
            display_name = f"{mode} {sense} - {filename}"
            self.file_listbox.insert(tk.END, display_name)
            self.update_plot()
            self.update_statistics()
            
        except Exception as e:
            # Handle file loading errors gracefully
            messagebox.showerror("Error", f"Failed to load CSV file:\n{file_path}\n\nError: {e}")
    
    def remove_selected_file(self):
        """
        Remove the currently selected curve from the comparison.
        Updates plots and statistics automatically.
        """
        selection = self.file_listbox.curselection()
        if selection:
            index = selection[0]
            self.loaded_curves.pop(index)
            self.file_listbox.delete(index)
            self.update_plot()
            self.update_statistics()
    
    def clear_all_files(self):
        """
        Remove all loaded curves from the comparison.
        Resets plots and statistics to empty state.
        """
        self.loaded_curves.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_plot()
        self.update_statistics()
    
    def update_plot(self):
        """
        Update the comparison plots with all currently loaded curves.
        Uses distinct colors, markers, and line styles for easy differentiation.
        """
        # Clear previous plot content
        self.ax1.clear()
        self.ax2.clear()
        
        # Handle empty curve list
        if not self.loaded_curves:
            self.ax1.text(0.5, 0.5, "No data loaded\nClick 'Browse Recent Measurements' to load your existing CSV files", 
                         ha='center', va='center', transform=self.ax1.transAxes)
            self.ax2.text(0.5, 0.5, "No data loaded", ha='center', va='center', transform=self.ax2.transAxes)
            self.canvas.draw()
            return
        
        # Define visual styling for curve differentiation
        colors = ['blue', 'red', 'green', 'purple', 'orange', 'brown', 'pink', 'gray', 'olive', 'cyan']
        markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p', '*', 'h', '+', 'x']
        linestyles = ['-', '--', '-.', ':', '-', '--', '-.', ':', '-', '--']
        
        # Plot I-V curves if enabled
        if self.show_iv_var.get():
            self.ax1.set_xlabel("Voltage (V)")
            self.ax1.set_ylabel("Current (A)")
            self.ax1.set_title("I-V Curve Comparison")
            self.ax1.grid(True, alpha=0.3)
            
            for i, curve in enumerate(self.loaded_curves):
                # Cycle through visual styles for each curve
                color = colors[i % len(colors)]
                marker = markers[i % len(markers)]
                linestyle = linestyles[i % len(linestyles)]
                label = f"{curve['mode']} {curve['sense']}"
                
                # Use absolute current values for consistent display
                current_abs = [abs(c) for c in curve['current']]
                
                # Plot with distinctive styling
                self.ax1.plot(curve['voltage'], current_abs, 
                             color=color, marker=marker, markersize=8, linewidth=3,
                             linestyle=linestyle, label=label, alpha=0.8, 
                             markeredgewidth=1, markeredgecolor='black')
            
            # Set axis limits to start from origin (like main application)
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
                # Use same styling scheme as I-V curves
                color = colors[i % len(colors)]
                marker = markers[i % len(markers)]
                linestyle = linestyles[i % len(linestyles)]
                label = f"{curve['mode']} {curve['sense']}"
                
                # Use absolute power values for consistent display
                power_abs = [abs(p) for p in curve['power']]
                
                # Plot power curve
                self.ax2.plot(curve['voltage'], power_abs, 
                             color=color, marker=marker, markersize=8, linewidth=3,
                             linestyle=linestyle, label=label, alpha=0.8,
                             markeredgewidth=1, markeredgecolor='black')
                
                # Highlight maximum power point with large star marker
                max_power_idx = power_abs.index(max(power_abs))
                max_power = power_abs[max_power_idx]
                max_power_voltage = curve['voltage'][max_power_idx]
                
                self.ax2.plot(max_power_voltage, max_power, 
                             color=color, marker='*', markersize=15, 
                             markeredgecolor='black', markeredgewidth=2)
            
            # Set axis limits to start from origin
            self.ax2.set_xlim(left=0)
            self.ax2.set_ylim(bottom=0)
            self.ax2.legend()
        
        # Finalize plot layout and display
        self.figure.tight_layout()
        self.canvas.draw()
    
    def update_statistics(self):
        """
        Calculate and display photovoltaic parameters for all loaded curves.
        Includes Pmp, Vmp, Imp, Voc, Isc, and fill factor calculations.
        """
        # Clear previous statistics
        self.stats_text.delete(1.0, tk.END)
        
        if not self.loaded_curves:
            self.stats_text.insert(tk.END, "No data loaded")
            return
        
        # Build statistics report
        stats_text = "Curve Statistics:\n" + "="*30 + "\n\n"
        
        for i, curve in enumerate(self.loaded_curves):
            stats_text += f"Curve {i+1}: {curve['mode']} {curve['sense']}\n"
            stats_text += f"File: {curve['filename']}\n"
            
            try:
                # Extract measurement arrays
                current_array = curve['current']
                voltage_array = curve['voltage']
                power_array = curve['power']
                
                # Convert to absolute values for consistent analysis
                current_abs = [abs(float(c)) for c in current_array]
                power_abs = [abs(float(p)) for p in power_array]
                voltage_vals = [float(v) for v in voltage_array]
                
                # Calculate maximum power point parameters
                max_power_idx = power_abs.index(max(power_abs))
                pmp = power_abs[max_power_idx]     # Maximum power
                vmp = voltage_vals[max_power_idx]  # Voltage at maximum power
                imp = current_abs[max_power_idx]   # Current at maximum power
                
                # Calculate open circuit voltage (Voc) - voltage at minimum current
                min_current_idx = current_abs.index(min(current_abs))
                voc = voltage_vals[min_current_idx]
                
                # Calculate short circuit current (Isc) - current at minimum voltage
                min_voltage_idx = voltage_vals.index(min(voltage_vals))
                isc = current_abs[min_voltage_idx]
                
                # Calculate fill factor: FF = (Pmp)/(Voc * Isc)
                if (voc * isc) > 0:
                    fill_factor = (pmp / (voc * isc)) * 100
                else:
                    fill_factor = 0
                
                # Add parameters to statistics text
                stats_text += f"Pmp: {pmp:.3f} W\n"      # Maximum power
                stats_text += f"Vmp: {vmp:.3f} V\n"      # Voltage at max power
                stats_text += f"Imp: {imp:.3f} A\n"      # Current at max power
                stats_text += f"Voc: {voc:.3f} V\n"      # Open circuit voltage
                stats_text += f"Isc: {isc:.3f} A\n"      # Short circuit current
                stats_text += f"FF: {fill_factor:.1f}%\n"  # Fill factor percentage
                
            except (ValueError, TypeError, IndexError) as e:
                # Handle calculation errors gracefully
                stats_text += f"Error calculating parameters: {e}\n"
            
            stats_text += "-"*25 + "\n\n"
        
        # Display complete statistics
        self.stats_text.insert(tk.END, stats_text)
    
    def export_comparison(self):
        """
        Export the comparison plot to a high-resolution image file.
        Supports PNG and PDF formats with timestamp in filename.
        """
        if not self.loaded_curves:
            messagebox.showwarning("Warning", "No data to export")
            return
        
        # Generate timestamp for unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Open save dialog
        plot_path = filedialog.asksaveasfilename(
            title="Save Comparison Plot",
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("PDF files", "*.pdf")],
            initialfile=f"IV_Comparison_{timestamp}.png"
        )
        
        if plot_path:
            try:
                # Save plot with high resolution and tight bounding box
                self.figure.savefig(plot_path, dpi=300, bbox_inches='tight')
                messagebox.showinfo("Export", f"Plot saved to:\n{plot_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save plot:\n{e}")

# Application entry point
if __name__ == "__main__":
    """
    Main execution block - creates and runs the GUI application.
    Only executes when script is run directly (not imported).
    """
    # Create main Tkinter window
    root = tk.Tk()
    root.geometry("950x850")      # Set initial window size
    root.resizable(True, True)    # Allow window resizing
    
    # Create and run the application
    app = IVAppCC(root)
    root.mainloop()  # Start GUI event loop