# cerca_aziende.py
import os
import re
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Carica .env
load_dotenv()

API_KEY = os.getenv("GOOGLE_API_KEY")
CSE_ID = os.getenv("GOOGLE_CSE_ID")

if not API_KEY or not CSE_ID:
    raise SystemExit("Errore: imposta GOOGLE_API_KEY e GOOGLE_CSE_ID nel file .env")

INPUT_FILE = "aziende.xlsx"
OUTPUT_FILE = "aziende_risultati.xlsx"

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; aziende-bot/1.0)"}

def google_search(query, api_key, cse_id, num=1):
    url = "https://www.googleapis.com/customsearch/v1"
    params = {
        'q': query,
        'key': api_key,
        'cx': cse_id,
        'num': num
    }
    r = requests.get(url, params=params, timeout=15)
    r.raise_for_status()
    data = r.json()
    return data.get("items", [])

def estrai_email(testo):
    match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', testo)
    return match.group(0) if match else ""

def truncate_to_50_words(text):
    words = re.split(r'\s+', text.strip())
    if len(words) <= 50:
        return " ".join(words)
    return " ".join(words[:50]) + "..."

def analizza_pagina(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=10)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")

        # Prova meta description
        descrizione = ""
        desc_tag = soup.find("meta", attrs={"name":"description"})
        if desc_tag and desc_tag.get("content"):
            descrizione = desc_tag["content"]

        # Se non trovi meta description prendi il primo <p> utile
        if not descrizione:
            p = soup.find("p")
            if p:
                descrizione = p.get_text(separator=" ", strip=True)

        # email nel testo della pagina
        email = estrai_email(r.text)

        return descrizione.strip(), email
    except Exception:
        return "", ""

def guess_categoria(snippet):
    s = snippet.lower()
    if "s.r.l" in s or "spa" in s or "srl" in s or "azienda" in s:
        return "Impresa"
    if "ristorante" in s or "pizzeria" in s or "chef" in s:
        return "Ristorazione"
    if "hotel" in s or "albergo" in s:
        return "Turismo / Ospitalità"
    if "software" in s or "it" in s or "sviluppo" in s:
        return "Tecnologia / Software"
    if "studio" in s or "avvocato" in s or "commercialista" in s:
        return "Professionisti / Studi"
    return ""

# Carica Excel (colonna A = "Nome Azienda")
df = pd.read_excel(INPUT_FILE, usecols=[0], header=0)
# Assicuriamoci che la colonna si chiami "Nome Azienda"
if df.columns[0].lower() not in ["nome azienda", "nome", "azienda"]:
    df.columns = ["Nome Azienda"]

df["Categoria Merceologica"] = ""
df["Descrizione"] = ""
df["Dominio"] = ""
df["Email"] = ""

for idx, nome in enumerate(df["Nome Azienda"].astype(str)):
    if not nome.strip():
        continue
    print(f"[{idx+1}/{len(df)}] Ricerca: {nome}")
    query = f'{nome} azienda sito ufficiale "azienda" "contatti"'
    try:
        results = google_search(query, API_KEY, CSE_ID, num=1)
    except Exception as e:
        print("Errore nella Google API:", e)
        time.sleep(2)
        continue

    if not results:
        print("Nessun risultato per:", nome)
        continue

    item = results[0]
    snippet = item.get("snippet", "")
    link = item.get("link", "")

    categoria = guess_categoria(snippet)
    descrizione_raw, email = analizza_pagina(link)

    # Se non ho descrizione dal sito uso lo snippet
    descrizione = descrizione_raw or snippet
    descrizione = truncate_to_50_words(descrizione)

    dominio_search = re.findall(r"https?://([A-Za-z0-9\.-]+)", link)
    dominio = dominio_search[0] if dominio_search else ""

    df.at[idx, "Categoria Merceologica"] = categoria
    df.at[idx, "Descrizione"] = descrizione
    df.at[idx, "Dominio"] = dominio
    df.at[idx, "Email"] = email

    time.sleep(1.0)  # rispetta le rate limits

df.to_excel(OUTPUT_FILE, index=False)
print("Fatto. File salvato:", OUTPUT_FILE)
