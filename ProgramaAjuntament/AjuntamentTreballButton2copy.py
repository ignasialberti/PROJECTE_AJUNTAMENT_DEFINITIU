import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import os
#SOLUCIÓ 2 ; SEMBLANT A LA ORIGINAL PERÒ AMB LA DIFERÈNCIA QUE BUSCA PARAULES ESPECIFIQUES EN ELS PDFS I NOMÉS ELS DESCARREGA SI LES CONTÉ
#Per a funcionar necessita que el procès sigui en pantalla completa, ja que si no no pot clicar en els botons de descàrrega dels PDFs.
# Crear una carpeta per guardar els documents descarregats
carpeta_descargas = "Documents_Descarregats_Ajuntament"
if not os.path.exists(carpeta_descargas):
    os.makedirs(carpeta_descargas)

# Configuració del navegador per descarregar fitxers automàticament
from selenium.webdriver.chrome.options import Options
options = Options()
prefs = {
    "download.default_directory": os.path.abspath(carpeta_descargas),  # Carpeta de descàrrega
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True  # Evita obrir PDFs al navegador
}
options.add_experimental_option("prefs", prefs)

# Configura el webdriver
driver = webdriver.Chrome(options=options)

# Llegir l'arxiu Excel
archivo_excel = "CPVajuntamentEDITATS.xlsx"
hoja_datos = pd.read_excel(archivo_excel)

# Mostrar els primers registres
print(hoja_datos.head())

# Comprovar que existeix la columna 'ENLLAC_PUBLICACIO'
if 'ENLLAC_PUBLICACIO' in hoja_datos.columns:
    # Extreure els enllaços de la columna
    enllacos = hoja_datos['ENLLAC_PUBLICACIO']

    print("Enllaços trobats:")
    for enllac in enllacos:
        print(enllac)
        try:
            # 1. Obre l'enllaç
            driver.get(enllac)

            # 2. Accepta cookies (si cal)
            try:
                accept_cookies_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Accepta')]"))
                )
                accept_cookies_button.click()
            except Exception:
                print("No s'ha trobat el botó de cookies. Continuant...")

            # 3. Cerca l'enllaç 'Anunci de licitació'
            try:
                Anunci_link = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.LINK_TEXT, "Anunci de licitació"))
                )
                Anunci_link.click()
                time.sleep(3)
                # AQUESTA COMANDA MOU EL CURSOR A LA PART DE BAIX DE LA PÀGINA; AIXI ES CARREGUEN TOTS ELS ELEMENTS (INDISPENSABLE)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)  # Esperem uns segons perquè es carreguin els elements

                # 4. Cerca i descarrega PDFs rellevants (com a botons)
                try:
                    pdf_buttons = driver.find_elements(By.XPATH, "//button[contains(text(),'.pdf')]")
                    for pdf_button in pdf_buttons:
                        pdf_text = pdf_button.text.strip().lower()

                        # Filtres: descarregar només PCAP, PPT o Plecs tècnics
                        if any(keyword in pdf_text for keyword in ["pcap","PCAP"," PCAP ","pca","PCA","_PCAP_","Administratiu"," PPT ","PPT ","plec administratiu", "ppt", "tècnic","PPT","Plec tècnic","plec tècnic",' PCAP']):
                            print(f"Descarregant PDF rellevant: {pdf_text}")

                            # Simula un clic al botó
                            pdf_button.click()
                            time.sleep(3)  # Espera perquè la descàrrega es completi

                            print(f"S'ha intentat descarregar el PDF rellevant amb el botó: {pdf_text}")
                        else:
                            print(f"Document ignorat: {pdf_text}")
                    pdf_links = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")

                except Exception as e:
                    print(f"Error en localitzar o clicar els botons de descàrrega de PDF: {e}")

            except Exception:
                print(f"No s'ha trobat l'enllaç 'Anunci de licitació' en l'enllaç: {enllac}. Continuant...")

        except Exception as e:
            print(f"Error en l'accés a l'enllaç {enllac}: {e}. Intentant el següent...")

else:
    print("La columna 'ENLLAC_PUBLICACIO' no existeix a l'arxiu.")

# Tancar el navegador
driver.quit()
