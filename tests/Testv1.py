import pyvisa

def test_visa_connection():
    try:
        rm = pyvisa.ResourceManager()
        resources = rm.list_resources()
        print("Detected instruments:", resources)

        if not resources:
            print("❌ No instrument detected.")
            return

        # You can replace this line with a specific address (e.g., "USB0::0x0AAD::0x0054::123456::INSTR")
        instrument_address = resources[0]
        print(f"Connecting to {instrument_address}...")

        instrument = rm.open_resource(instrument_address)
        instrument.timeout = 5000  # 5 seconds

        # *IDN? command to identify the device
        idn = instrument.query("*IDN?")
        print("✅ *IDN? response:", idn.strip())

        # Read remote/local 2W/4W mode if supported
        try:
            sense_mode = instrument.query("SYST:SENS:REM?")
            mode_str = "4-Wire" if sense_mode.strip() == '1' else "2-Wire"
            print(f"🔍 Current measurement mode: {mode_str}")
        except Exception as e:
            print("⚠️ Unable to read measurement mode:", e)

        instrument.close()
        print("✅ Connection tested successfully.")

    except Exception as e:
        print("❌ Error during connection test:", e)

if __name__ == "__main__":
    test_visa_connection()
