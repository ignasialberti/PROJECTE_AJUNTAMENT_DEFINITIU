import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import os
import requests

# Llegir l'arxiu Excel
archivo_excel = "CPVajuntamentEDITATS.xlsx"
hoja_datos = pd.read_excel(archivo_excel)

# Comprovar que existeixen les columnes 'CODI_EXPEDIENT' i 'ENLLAC_PUBLICACIO'
if 'CODI_EXPEDIENT' in hoja_datos.columns and 'ENLLAC_PUBLICACIO' in hoja_datos.columns:
    # Extreure els valors de les columnes
    codis_expedient = hoja_datos['CODI_EXPEDIENT']
    enllacos = hoja_datos['ENLLAC_PUBLICACIO']

    for codi_expedient, enllac in zip(codis_expedient, enllacos):
        # Crear una carpeta per guardar els documents descarregats
        carpeta_descargas = os.path.join("Documents_Descarregats_Ajuntament", str(codi_expedient))
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

            # AQUESTA COMANDA MOU EL CURSOR A UNA PART DE LA PÀGINA; AIXI ES CARREGUEN ELS ELEMENTS NECESSARIS
            try:
                Anunci_link = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.LINK_TEXT, "Anunci de licitació"))
                )
                Anunci_link.click()
                time.sleep(3)
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight * 0.7);")  # Baixa fins al 70% de l'alçada total
                time.sleep(2)  # Esperem uns segons perquè es carreguin els elements
            except Exception as e:
                print(f"Error durant la interacció amb l'anunci de licitació: {e}")

            # 4.1 Buscar botons amb text que contingui '.pdf'
            try:
                pdf_buttons = driver.find_elements(By.XPATH, "//button[contains(text(),'.pdf')]")
                for pdf_button in pdf_buttons:
                    pdf_text = pdf_button.text.strip().lower()

                    # Filtres: descarregar només PCAP, PPT o Plecs tècnics
                    if any(keyword in pdf_text for keyword in ["pcap", "PTT", "condicions", "plec", "PCT", "PCAP", "PPT", "plec administratiu", "plec tècnic", "memòria", "memoria", "tècnic"]):
                        print(f"Descarregant PDF rellevant: {pdf_text}")

                        # Simula un clic al botó
                        pdf_button.click()
                        time.sleep(3)  # Espera perquè la descàrrega es completi
                        print(f"S'ha intentat descarregar el PDF rellevant amb el botó: {pdf_text}")
                    else:
                        print(f"Document ignorat (botó): {pdf_text}")
            except Exception as e:
                print(f"Error durant la cerca i descàrrega de PDFs amb botons: {e}")

            # 4.2 Buscar enllaços amb qualsevol href
            try:
                pdf_links = driver.find_elements(By.XPATH, "//a[@href]")
                print(f"S'han trobat {len(pdf_links)} enllaços amb href.")
                
                for pdf_link in pdf_links:
                    pdf_text = pdf_link.text.strip().lower()
                    print(f"Enllaç trobat: {pdf_text}")

                    if any(keyword in pdf_text for keyword in ["pcap", "pca"]):
                        pdf_url = pdf_link.get_attribute("href")
                        print(f"Descarregant PDF rellevant des de l'URL: {pdf_url}")

                        try:
                            # Intentem descarregar el PDF amb Selenium
                            driver.get(pdf_url)
                            print(f"S'ha iniciat la descàrrega del fitxer amb Selenium: {pdf_url}")
                            time.sleep(5)  # Esperem perquè la descàrrega es completi
                        except Exception as selenium_error:
                            print(f"Error durant la descàrrega amb Selenium: {selenium_error}")
                        
                        # Comprovem si es pot utilitzar requests com a pla B
                        try:
                            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"}
                            response = requests.get(pdf_url, headers=headers)
                            if response.status_code == 200:
                                filename = os.path.join(carpeta_descargas, pdf_url.split('/')[-1])
                                with open(filename, 'wb') as f:
                                    f.write(response.content)
                                print(f"Descarregat correctament amb requests: {filename}")
                            else:
                                print(f"No s'ha pogut descarregar el fitxer amb requests. Codi d'estat: {response.status_code}")
                        except Exception as requests_error:
                            print(f"Error durant la descàrrega amb requests: {requests_error}")
                    else:
                        print(f"Document ignorat: {pdf_text}")
            except Exception as e:
                print(f"Error durant la descàrrega dels enllaços PDF: {e}")

        finally:
            # Tancar el navegador
            driver.quit()
else:
    print("Les columnes 'CODI_EXPEDIENT' o 'ENLLAC_PUBLICACIO' no existeixen a l'arxiu.")