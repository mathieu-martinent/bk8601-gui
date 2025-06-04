import pyvisa
import time

def simple_test():
    rm = pyvisa.ResourceManager()
    resources = rm.list_resources()
    print("Instruments détectés :", resources)

    if resources:
        instrument_address = resources[0]
        print(f"Connexion à {instrument_address}...")

        try:
            instrument = rm.open_resource(instrument_address)
            instrument.timeout = 0  # 5 secondes
            instrument.write("*RST")
            time.sleep(1)  # Attendre un peu après la réinitialisation
            idn = instrument.query("*IDN?")
            print(f"Réponse *IDN? : {idn.strip()}")
            instrument.close()
        except pyvisa.VisaIOError as e:
            print(f"Erreur de communication : {e}")
    else:
        print("Aucun instrument détecté.")

if __name__ == "__main__":
    simple_test()
