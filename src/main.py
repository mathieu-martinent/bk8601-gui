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

class IVAppCC:
    def __init__(self, root):
        self.root = root
        self.root.title("I-V Curve Measurement (CC/CV Mode)")

        # VISA instrument selection
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

        # Progress bar to show sweep progress
        self.progress = ttk.Progressbar(root, orient='horizontal', mode='determinate')
        self.progress.grid(row=12, column=0, columnspan=3, sticky="ew", padx=5)

        # Plot area for displaying the I-V curve
        self.figure = plt.Figure(figsize=(7, 6), dpi=100)
        self.ax = self.figure.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.figure, master=root)
        self.canvas.get_tk_widget().grid(row=13, column=0, columnspan=3, sticky="nsew")

        # Configure resizing behavior
        root.grid_rowconfigure(13, weight=1)
        root.grid_columnconfigure(1, weight=1)

        # Flags for sweep state
        self.stop_requested = False
        self.sweep_running = False

        # Bind Enter and Escape keys
        self.root.bind('<Return>', self.on_enter)
        self.root.bind('<Escape>', self.on_escape)

    def choose_output_dir(self):
        new_dir = filedialog.askdirectory(initialdir=os.getcwd(), title="Choose output directory")
        if new_dir:
            self.output_dir = new_dir
            self.folder_label.config(text=f"Output directory: {self.output_dir}")

    def on_enter(self, event):
        if not self.sweep_running:
            self.start_sweep_thread()

    def on_escape(self, event):
        if self.sweep_running:
            self.request_stop()

    def start_sweep_thread(self):
        if self.sweep_running:
            return  # Prevent double start
        self.sweep_running = True
        thread = threading.Thread(target=self.start_sweep)
        thread.daemon = True
        thread.start()

    def request_stop(self):
        self.stop_requested = True

    def update_labels(self, *args):
        selected_mode = self.mode_var.get()
        if selected_mode == "CC":
            self.start_label.config(text="Start Current (A):")
            self.end_label.config(text="End Current (A):")
            self.step_label.config(text="Step Current (A):")
        else:
            self.start_label.config(text="Start Voltage (V):")
            self.end_label.config(text="End Voltage (V):")
            self.step_label.config(text="Step Voltage (V):")

    def start_sweep(self):
        self.stop_requested = False  # Reset stop flag at each sweep start
        try:
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
            if instrument_address == "Simulated Instrument":
                load = self.create_simulated_instrument()
            else:
                load = self.rm.open_resource(instrument_address)
                load.timeout = 5000
                load.write("*CLS")

            selected_mode = self.mode_var.get()
            mode_mapping = {"CC": "CURR", "CV": "VOLT"}
            if load.query("FUNC?").strip() != mode_mapping[selected_mode]:
                load.write(f"FUNC {mode_mapping[selected_mode]}")

            # Move INPUT ON after mode and protection configuration
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

            sense_command = "REM:SENS ON" if self.sense_mode_var.get() == "4-Wire" else "REM:SENS OFF"
            load.write(sense_command)
            time.sleep(0.5)

            # Enable the input only after all configuration is done
            load.write("INPUT ON")

            self.ax.clear()
            if hasattr(self, 'ax2'):
                self.figure.delaxes(self.ax2)
            self.ax2 = self.ax.twinx()

            # Always set axis labels right after clearing
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

            currents, voltages, powers = [], [], []
            current = i_start
            step = i_step if i_end >= i_start else -i_step
            total_steps = int(abs((i_end - i_start) / i_step)) + 1
            self.progress["maximum"] = total_steps
            self.progress["value"] = 0

            # --- Force starting setpoint and wait before sweep ---
            if selected_mode == "CC":
                load.write(f"CURR {i_start:.3f}")
            else:
                load.write(f"VOLT {i_start:.3f}")
            time.sleep(sleep_time)  # Let the load stabilize at the starting point

            # Démarre la boucle SANS ajouter de point avant
            for count in range(total_steps):
                if self.stop_requested:
                    messagebox.showinfo("Sweep Stopped", "Sweep was stopped by the user.")
                    break
                try:
                    if selected_mode == "CC":
                        load.write(f"CURR {current:.3f}")
                    else:
                        load.write(f"VOLT {current:.3f}")
                    time.sleep(sleep_time)  # Let the load stabilize

                    voltage = float(load.query("MEAS:VOLT?"))
                    actual_current = float(load.query("MEAS:CURR?"))
                    print(f"Measured: V={voltage}, I={actual_current}")
                    power = voltage * actual_current

                    if voltage_limit is not None and voltage > voltage_limit:
                        raise Exception("Voltage exceeded protection limit.")
                    if current_limit is not None and actual_current > current_limit:
                        raise Exception("Current exceeded protection limit.")

                    EPS = 1e-4  # tolérance pour éviter les doublons dus à l'arrondi

                    if len(currents) == 0 or abs(actual_current - currents[-1]) > EPS or abs(voltage - voltages[-1]) > EPS:
                        currents.append(actual_current)
                        voltages.append(voltage)
                        powers.append(power)

                    if hasattr(self, 'line_iv'):
                        self.line_iv.remove()
                        del self.line_iv
                    self.line_iv, = self.ax.plot(voltages, currents, label="I-V Curve", color='blue')

                    if hasattr(self, 'line_power'):
                        self.line_power.remove()
                        del self.line_power
                    self.line_power, = self.ax2.plot(voltages, powers, label="P-V Curve", color='red')

                    self.ax.relim()
                    self.ax.autoscale_view()
                    self.ax2.relim()
                    self.ax2.autoscale_view()

                    self.canvas.draw()
                    self.root.update_idletasks()
                    self.progress["value"] = count + 1
                except Exception as e:
                    messagebox.showwarning("Protection Triggered", f"Sweep stopped: {e}")
                    break
                current += step

            load.write("INPUT OFF")
            load.close()

            # Swap axes: Voltage (V) on X, Current (A) on Y
            self.line_iv.set_data(voltages, currents)
            self.line_power.set_data(voltages, powers)
            self.ax.set_xlabel("Voltage (V)")
            self.ax.set_ylabel("Current (A)", color='b')
            self.ax2.set_ylabel("Power (W)", color='r')
            self.ax.relim()
            self.ax.autoscale_view()
            self.ax2.relim()
            self.ax2.autoscale_view()
            self.canvas.draw()

            if powers:
                pmp = max(powers)
                idx = powers.index(pmp)
                vmp = voltages[idx]
                imp = currents[idx]

                self.pmp_point, = self.ax2.plot(vmp, pmp, 'ro')
                self.canvas.draw()

                self.vmp_point, = self.ax.plot(vmp, imp, 'go', markersize=8, label="Vmp/Imp")
                self.vmp_annotation = self.ax.annotate(
                    "Vmp/Imp",
                    (vmp, imp),
                    textcoords="offset points",
                    xytext=(-60, 10),
                    ha='right',
                    color='g',
                    fontsize=10,
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="green", lw=1)
                )
                self.canvas.draw()

                self.pmp_annotation = self.ax2.annotate(
                    "Pmp",
                    (vmp, pmp),
                    textcoords="offset points",
                    xytext=(0, 10),
                    ha='center',
                    color='r',
                    fontsize=10,
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="red", lw=1)
                )

                self.vmp_annotation = self.ax.annotate(
                    "Vmp/Imp",
                    (vmp, imp),
                    textcoords="offset points",
                    xytext=(-60, 10),
                    ha='right',
                    color='g',
                    fontsize=10,
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="green", lw=1)
                )
                self.canvas.draw()
            else:
                pmp = vmp = imp = None

            if hasattr(self, 'summary_annotation'):
                try:
                    self.summary_annotation.remove()
                except Exception:
                    pass
                del self.summary_annotation

            summary_text = f"Pmp = {pmp:.2f} W   Vmp = {vmp:.2f} V   Imp = {imp:.2f} A"
            self.summary_annotation = self.ax.annotate(
                summary_text,
                xy=(0.5, 1.08), xycoords='axes fraction',
                ha='center', va='bottom',
                fontsize=14, color='purple',
                bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="purple", lw=2)
            )
            self.canvas.draw()

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

            if self.save_csv_var.get():
                csv_path = os.path.join(self.output_dir, f"{base_filename}.csv")
                with open(csv_path, mode='w', newline='') as file:
                    writer = csv.writer(file)
                    writer.writerow(["Current (A)", "Voltage (V)", "Power (W)"])
                    for i in range(len(currents)):
                        writer.writerow([currents[i], voltages[i], powers[i]])
                    writer.writerow([])  # Ligne vide
                    writer.writerow(["Parameter", "Value"])
                    for param, value in params:
                        writer.writerow([param, value])
                print(f"Data saved to {csv_path}")

            if self.save_png_var.get():
                png_path = os.path.join(self.output_dir, f"{base_filename}.png")
                self.figure.savefig(png_path)
                print(f"Plot saved to {png_path}")

            message = f"Sweep completed.\nPmp = {pmp:.2f} W\nVmp = {vmp:.2f} V\nImp = {imp:.2f} A" if pmp else "Sweep completed with no power data."
            messagebox.showinfo("Sweep Complete", message)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{e}")

        finally:
            self.sweep_running = False
            self.start_button.config(state='normal')

    def create_simulated_instrument(self):
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

if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("950x850")
    root.resizable(True, True)
    app = IVAppCC(root)
    root.mainloop()