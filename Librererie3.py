# Importiamo le librerie necessarie
import subprocess
import sys
import os
subprocess.call("clear" if os.name == "posix" else "cls", shell=True)

# Una lista di tuple che contengono il nome di una libreria e le sue dipendenze
libraries = [("art", []), 
             ("scapy", []), 
             ("python-nmap", []), 
             ("wakeonlan", []), 
             ("openpyxl", []),
             ("schedule", []),
             ("time", []),  
             ("socket", []),           
             ("ipaddress", []), 
             ("pynput", []),                                                              
             ("pandas", [])]
# Una lista che tiene traccia delle librerie che sono state installate
installed_libraries = []

# Verifica se una libreria è installata e, in caso contrario, la installa
def check_library(library, dependencies):
    try:
        __import__(library)
        installed_libraries.append(library)
    except ImportError:
        subprocess.call([sys.executable, "-m", "pip", "install", library])
        installed_libraries.append(library)

# Verifica l'installazione di tutte le librerie
for library, dependencies in libraries:
    check_library(library, dependencies)

# Importiamo le librerie solo dopo che sono state installate
import getpass
import tqdm
from tqdm import tqdm
import art
import scapy.all as scapy
import nmap
import wakeonlan
import openpyxl
import time
import schedule
import pandas as pd
import ipaddress
import socket
import logging

def export_to_excel(hosts):

    # Crea una lista di tuple che includono nome del dispositivo, indirizzo IP e indirizzo MAC
    host_info = []
    for host in hosts:
        if len(host) == 2:
            host_info.append((host[0], host[1], "N/A"))
        elif len(host) == 3:
            host_info.append((host[0], host[1], host[2]))
    
    try:
        # Crea un dataframe Pandas a partire dalla lista di informazioni sul host
        df = pd.DataFrame(host_info, columns=["Nome dispositivo", "Indirizzo IP", "Indirizzo MAC"])

        # Esporta il dataframe in un file Excel con nome "risultati_scansione.xlsx"
        # e imposta l'indice a False in modo da non includere una colonna di indice nella cartella di lavoro
        df.to_excel("risultati_scansione.xlsx", index=False)
    except Exception as e:
        print("Errore durante l'esportazione dei risultati in Excel:", e)

def show_device_details(hosts):
    print("Seleziona un dispositivo:")
    for i, host in enumerate(hosts):
        print("{}. {}: {}".format(i + 1, host[0], host[1]))

    selected_index = input("Inserisci il numero del dispositivo selezionato: ")
    try:
        selected_index = int(selected_index) - 1
        if selected_index < 0 or selected_index >= len(hosts):
            raise ValueError
    except ValueError:
        print("Numero non valido.")
        return
    
    selected_host = hosts[selected_index]
    print("Nome dispositivo: {}".format(selected_host[0]))
    print("Indirizzo IP: {}".format(selected_host[1]))
    # Potresti utilizzare una libreria come 'os-fingerprint' per ottenere ulteriori informazioni sul dispositivo

logging.basicConfig(filename="network_monitor.log", level=logging.INFO,
                    format="%(asctime)s: %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

def real_time_monitoring(hosts):
    logging.info("Inizio del monitoraggio in tempo reale...")
    print("Inizio del monitoraggio in tempo reale...")
    print("Premi 'q' + invio per terminare il monitoraggio.")

    while True:
        new_hosts = network_scan()
        added_hosts = [host for host in new_hosts if host not in hosts]
        removed_hosts = [host for host in hosts if host not in new_hosts]

        if added_hosts:
            logging.info("Dispositivi connessi:")
            print("Dispositivi connessi:")
            for host in added_hosts:
                logging.info("{}: {}".format(host[0], host[1]))
                print("{}: {}".format(host[0], host[1]))
        
        if removed_hosts:
            logging.info("Dispositivi disconnessi:")
            print("Dispositivi disconnessi:")
            for host in removed_hosts:
                logging.info("{}: {}".format(host[0], host[1]))
                print("{}: {}".format(host[0], host[1]))

        hosts = new_hosts
        time.sleep(5)

        user_input = input()
        if user_input.lower() == "q":
            logging.info("Monitoraggio terminato.")
            print("Monitoraggio terminato.")
            break

def select_subnet():
    print("Seleziona una sottorete:")
    subnet = input("Inserisci un indirizzo IP/maschera di sottorete (es. 192.168.1.0/24): ")
    
    try:
        subnet = ipaddress.ip_network(subnet)
        hosts = []
        for host in subnet.hosts():
            try:
                hostname = socket.gethostbyaddr(str(host))[0]
                hosts.append((hostname, str(host)))
            except socket.herror:
                hosts.append(("", str(host)))
        
        return hosts
    except ValueError:
        print("Indirizzo IP/maschera di sottorete non valido.")
        return []

def get_schedule_data():
    schedule_data = []
    schedule_file = "schedule.txt"
    try:
        with open(schedule_file, "r") as file:
            for line in file:
                device_name, mac_address, on_time, off_time = line.strip().split(",")
                schedule_data.append((device_name, mac_address, on_time, off_time))
    except FileNotFoundError:
        return []
    except Exception as e:
        raise Exception("Errore durante la lettura dei dati di pianificazione: {}".format(str(e)))
    return schedule_data

def export_excel(file_path, schedule_data):
    if len(schedule_data) == 0:
        # Creazione di un modello vuoto
        data = {"Dispositivo": [], "Indirizzo MAC": [], "Orario accensione": [], "Orario spegnimento": []}
        df = pd.DataFrame(data)
    else:
        # Scrittura dei dati di pianificazione
        data = {"Dispositivo": [d[0] for d in schedule_data], 
                "Indirizzo MAC": [d[1] for d in schedule_data], 
                "Orario accensione": [d[2] for d in schedule_data], 
                "Orario spegnimento": [d[3] for d in schedule_data]}
        df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)

# Stampa l'intestazione del menu con arte ASCII
def text_art_menu():
    print("\033c", end="") # pulisce lo schermo del terminale
    print("\033[1;32;40m", end="") # imposta il colore del testo in verde
    print("\033[1m", end="") # imposta il testo in grassetto
    print("\033[37m", end="") # imposta il colore del testo in bianco
    art.tprint("Turn On Manager", "georgia11")
    print("\n")
    print("Powered for", end="")
    print("\n")
    art.tprint("\t\tPiazza HR", "lineblocks")
    print("\033[0m", end="") # ripristina il colore e lo stile del testo di default
    print("\n")

# Stampa una barra di caricamento
def loading_bar():
    with tqdm(total=100, desc="Caricamento...", leave=False) as pbar:
        for i in range(100):
            time.sleep(0.02)
            pbar.update(1)

def network_scan():
    print("Scansione in corso...")
    nm = nmap.PortScanner()
    nm.scan('127.0.0.1', '22-443')
    hosts = [(x, nm[x]['status']['state']) for x in nm.all_hosts()]
    return hosts
# Chiamiamo la funzione per stampare l'intestazione del menu con arte ASCII
text_art_menu()
# Chiamiamo la funzione per stampare la barra di caricamento
loading_bar()

def configure_schedule(device_name):
    schedule = {}
    days_of_week = ["lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato", "domenica"]
    for day in days_of_week:
        on_time = input("Inserisci l'ora di accensione per {} (formato HH:MM): ".format(day))
        off_time = input("Inserisci l'ora di spegnimento per {} (formato HH:MM): ".format(day))
        try:
            on_time = datetime.datetime.strptime(on_time, "%H:%M").time()
            off_time = datetime.datetime.strptime(off_time, "%H:%M").time()
        except ValueError:
            print("Errore: Il formato dell'ora non è valido.")
            sys.exit()
        schedule[day] = (on_time, off_time)
    return schedule

def import_excel(file_path):
    try:
        workbook = openpyxl.load_workbook(file_path)
    except Exception as e:
        raise Exception("Errore durante l'importazione del file: {}".format(str(e)))
    sheet = workbook.active
    schedule_data = []
    for row in sheet.iter_rows(values_only=True):
        device_name = row[0]
        mac_address = row[1]
        on_time = row[2]
        off_time = row[3]
        schedule_data.append((device_name, mac_address, on_time, off_time))
    return schedule_data

def send_signal(mac_address, signal):
    if signal == "wake":
        result = subprocess.call(["wakeonlan", mac_address])
        if result == 0:
            print("Segnale Wake-on-LAN inviato con successo a: " + mac_address)
        else:
            print("Errore durante l'invio del segnale Wake-on-LAN a: " + mac_address)
    elif signal == "sleep":
        print("Segnale Sleep-on-LAN non implementato.")
    else:
        print("Segnale non valido.")

# Menu principale
while True:
    subprocess.run("cls", shell=True)
    text_art_menu()
    print("\nMenu:")
    print("1. Scansione e Monitoraggio")
    print("2. Gestione e Routine")
    print("3. Accendi tutto")
    scelta = input("Seleziona un'opzione (1, 2 o 3): ")

    if not scelta.isdigit():
        print("Selezione non valida. Inserire solo numeri interi.")
        continue
    
    
    scelta = int(scelta)
    os.system("clear")
    if scelta not in [1, 2, 3]:
        print("Selezione non valida. Inserire solo i numeri 1, 2 o 3.")
        continue
    if scelta == 1:
        # Scansione e monitoraggio della rete
        os.system("clear")
        print("Hai selezionato l'opzione 1: Scansione e Monitoraggio della Rete")
        hosts = []
        while True:
            os.system("clear")
            print("Sottomenu:")
            print("1. Scansione")
            print("2. Visualizzazione dei dispositivi")
            print("3. Configurazione")
            print("4. Torna al menu principale")
            sottomenu = input("Scegli un'opzione: ")
    
            if sottomenu == "1":
                os.system("clear")
                print("Sottomenu Scansione:")
                print("1. Esegui scansione")
                print("2. Monitoraggio in tempo reale")
                print("3. Selezione della sottorete")
                sottomenu_scansione = input("Scegli un'opzione: ")
    
                if sottomenu_scansione == "1":
                    hosts = network_scan()
                    if hosts:
                        print("Dispositivi connessi alla rete:")
                        for host in hosts:
                            print("{}: {}".format(host[0], host[1]))
                    else:
                        print("Nessun dispositivo connesso alla rete.")
                    input("Premi invio per continuare...")
                elif sottomenu_scansione == "2":
                    # Monitoraggio in tempo reale
                    real_time_monitoring(hosts)
                elif sottomenu_scansione == "3":
                    # Selezione della sottorete
                    hosts = select_subnet()
                else:
                    print("Opzione non valida. Riprova.")
            elif sottomenu == "2":
                os.system("clear")
                print("Sottomenu Visualizzazione dei dispositivi:")
                print("1. Visualizzazione dettagliata")
                print("2. Esporta in Excel")
                sottomenu_dispositivi = input("Scegli un'opzione: ")
    
                if sottomenu_dispositivi == "1":
                    # Visualizzazione dettagliata dei dispositivi
                    show_device_details(hosts)
                elif sottomenu_dispositivi == "2":
                    if hosts:
                        export_to_excel(hosts)
                        print("Esportazione completata con successo.")
                    else:
                        print("Esegui prima una scansione della rete.")
                    input("Premi invio per continuare...")
                else:
                    print("Opzione non valida. Riprova.") 
            if sottomenu == "3":
                os.system("clear")
                print("Sottomenu Configurazione:")
                print("1. Configurazione delle impostazioni")
                sottomenu_configurazione = input("Scegli un'opzione: ")
            
                if sottomenu_configurazione == "1":
                            # Configurazione delle impostazioni
                            configure_settings()
                else:
                            print("Opzione non valida. Riprova.")
            elif sottomenu == "4":
                break
            else:
                        print("Opzione non valida. Riprova.")
            

    elif scelta == 2:
        # Gestione e routine
        print("Hai selezionato l'opzione 2: Gestione e Routine")
        print("1. Accendi dispositivo")
        print("2. Spegni dispositivo")
        print("3. Importa lista da file Excel")
        print("4. Seleziona dispositivo recentemente utilizzato")
        print("5. Scansione rapida della rete")
        print("6. Configura routine")
        print("7. Esporta lista in un file Excel compilabile")


        # Chiedi all'utente di inserire la scelta
        scelta_2 = input("Seleziona un'opzione (1, 2, 3, 4, 5, 6, o 7): ")

        # Prova a convertire la scelta in un intero
        try:
            scelta_2 = int(scelta_2)
        except ValueError:
            # Se la conversione non riesce, stampa un messaggio di errore e termina l'esecuzione
            print("Errore: La scelta deve essere 1, 2, 3, 4, 5, 6, o 7.")
            sys.exit()  
                # Verifica che la scelta sia valida
        # Verifica che la scelta sia valida
        if scelta_2 not in [1, 2, 3, 4, 5, 6, 7]:
            # Se la scelta non è valida, stampa un messaggio di errore e termina l'esecuzione
            print("Errore: La scelta deve essere 1, 2, 3, 4, 5, 6, o 7.")
            sys.exit()
                
        # Esegui l'azione corrispondente alla scelta
        if scelta_2 == 1:
            # Accensione dispositivo
            print("Hai selezionato l'opzione 1: Accendi dispositivo")
            mac_address = input("Inserisci l'indirizzo MAC del dispositivo: ")
            send_signal(mac_address, "wake")
        elif scelta_2 == 2:
            # Spegnimento dispositivo
            print("Hai selezionato l'opzione 2: Spegni dispositivo")
            mac_address = input("Inserisci l'indirizzo MAC del dispositivo: ")
            send_signal(mac_address, "sleep")
        elif scelta_2 == 3:
            # Importazione lista da file Excel
            print("Hai selezionato l'opzione 3: Importa lista da file Excel")
            file_path = input("Inserisci il percorso del file: ")
            try:
                hosts = import_excel(file_path)
            except Exception as e:
                print("Si è verificato un errore durante l'importazione del file: {}".format(str(e)))
                sys.exit()
            for host in hosts:
                try:
                    send_signal(host[1], "wake")
                except Exception as e:
                    print("Si è verificato un errore durante l'invio del segnale per il dispositivo {}: {}".format(host[0], str(e)))
                    continue
        elif scelta_2 == 4:
            # Seleziona dispositivo recentemente utilizzato
            print("Hai selezionato l'opzione 4: Seleziona dispositivo recentemente utilizzato")
            recent_device = select_recent_device()
            if recent_device is None:
                print("Non ci sono dispositivi recentemente utilizzati.")
                sys.exit()
            try:
                send_signal(recent_device[1], "wake")
            except Exception as e:
                print("Si è verificato un errore durante l'invio del segnale per il dispositivo {}: {}".format(recent_device[0], str(e)))
                sys.exit()
        elif scelta_2 == 5:
            # Scansione rapida della rete
            print("Hai selezionato l'opzione 5: Scansione rapida della rete")
            hosts = quick_network_scan()
            if len(hosts) == 0:
                print("Non sono stati trovati dispositivi connessi.")
                sys.exit()
            print("I seguenti dispositivi sono stati trovati connessi:")
            for i, host in enumerate(hosts):
                print("{}. {} ({})".format(i + 1, host[0], host[1]))
            accendi_tutti = input("Vuoi accendere tutti i dispositivi? (s/n) ")
            if accendi_tutti.lower() == 's':
                for host in hosts:
                    try:
                        send_signal(host[1], "wake")
                    except Exception as e:
                        print("Si è verificato un errore durante l'invio del segnale per il dispositivo {}: {}".format(host[0], str(e)))
                        continue
        elif scelta_2 == 6:
            # Configura routine
            print("Hai selezionato l'opzione 6: Configura routine")
            hosts = network_scan()
            if len(hosts) == 0:
                print("Non sono stati trovati dispositivi connessi.")
                sys.exit()
            print("I seguenti dispositivi sono stati trovati connessi:")
            for i, host in enumerate(hosts):
                print("{}. {} ({})".format(i + 1, host[0], host[1]))
            device_index = input("Seleziona un dispositivo (numero): ")
            try:
                device_index = int(device_index) - 1
            except ValueError:
                print("Errore: La selezione deve essere un numero.")
                sys.exit()
            if device_index < 0 or device_index >= len(hosts):
                print("Errore: La selezione deve essere compresa tra 1 e {}.".format(len(hosts)))
                sys.exit()
            device = hosts[device_index]
            schedule = configure_schedule(device[0])
            save_schedule(device[0], device[1], schedule)
        elif scelta_2 == 7:
            # Importa o esporta lista in un file Excel compilabile
            print("Hai selezionato l'opzione 7: Importa o esporta lista da file Excel")
            print("1. Importa lista")
            print("2. Esporta lista")
            print("3. Scarica modello")
            sub_choice = input("Seleziona un'opzione (1, 2 o 3): ")
            if sub_choice == "1":
                file_path = input("Inserisci il percorso del file di origine: ")
                try:
                    schedule_data = import_excel(file_path)
                except Exception as e:
                    print("Errore durante l'importazione del file: {}".format(str(e)))
                    sys.exit()
                print("L'importazione del file è stata completata con successo.")
            elif sub_choice == "2":
                try:
                    schedule_data = get_schedule_data()
                    if len(schedule_data) == 0:
                        print("Non ci sono dati da esportare.")
                        sys.exit()
                    export_excel("schedule.xlsx", schedule_data)
                except Exception as e:
                    print("Errore durante l'esportazione del file: {}".format(str(e)))
                    sys.exit()
                print("L'esportazione del file è stata completata con successo.")
            elif sub_choice == "3":
                file_path = "schedule_template.xlsx"
                try:
                    export_excel(file_path, [])
                except Exception as e:
                    print("Errore durante il download del modello: {}".format(str(e)))
                    sys.exit()
                print("Il modello è stato scaricato con successo.")

        else:
            sys.exit()
        
        
    elif scelta == 3:
        # Stampa un messaggio che indica che l'utente ha selezionato l'opzione 3
        print("Hai selezionato l'opzione 3: Accendi tutti i dispositivi")
        
        # Esegui la scansione della rete per trovare i dispositivi
        hosts = network_scan()
        
        # Itera attraverso ogni dispositivo trovato
        for host in hosts:
            # Chiedi all'utente di inserire l'indirizzo MAC del dispositivo corrente
            mac_address = input("Inserisci l'indirizzo MAC del dispositivo " + host[0] + ": ")
            
            # Prova a inviare un segnale di accensione al dispositivo
            try:
                send_signal(mac_address, "wake")
            except Exception as e:
                # Se si verifica un errore, stampa un messaggio di errore
                print("Si è verificato un errore durante l'invio del segnale per il dispositivo {}: {}".format(host[0], str(e)))
                
                # Continua con l'iterazione successiva
                continue

    else:
        print("Selezione non valida. Inserire solo i numeri 1, 2 o 3.")