import pyvisa
import time
import csv

# Sweep configuration
start_current = 0.0       # Amps
end_current = 1.5         # Amps (adjust based on the cell's short-circuit current)
step_current = 0.05       # Amps per step
delay = 0.2               # Seconds between each step
output_csv = "iv_curve_cc.csv"

# Connect to the electronic load
rm = pyvisa.ResourceManager()
load = rm.open_resource("USB0::0x2EC7::0x8800::802199042787070066::INSTR")  

# Identify the instrument
print("Instrument ID:", load.query("*IDN?"))

# Configure the load in Constant Current mode
load.write("MODE CC")          # Set mode to Constant Current
load.write("INPUT ON")         # Enable the input

# Lists to store measurement data
currents = []
voltages = []

i = start_current
while i <= end_current:
    load.write(f"CURR {i:.3f}")               # Apply current setpoint
    time.sleep(delay)
    v = float(load.query("MEAS:VOLT?"))       # Measure voltage
    im = float(load.query("MEAS:CURR?"))      # Measure actual current
    print(f"I = {im:.3f} A, V = {v:.3f} V")
    currents.append(im)
    voltages.append(v)
    i += step_current

# Turn off the load input
load.write("INPUT OFF")

# Save the data to a CSV file
with open(output_csv, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["Current (A)", "Voltage (V)"])
    for i, v in zip(currents, voltages):
        writer.writerow([i, v])

print(f"I-V data saved to {output_csv}")
