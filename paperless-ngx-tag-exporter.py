#!/usr/bin/env python3

import os
import requests
import pandas as pd
import argparse
from configparser import ConfigParser
from tqdm import tqdm
import json
import sys
import inspect
from datetime import datetime
import shutil
from openpyxl.styles import Font, PatternFill
from openpyxl.formatting.rule import FormulaRule
import locale

# Setzt das locale-Systemformat, abhängig von Systemeinstellungen
locale.setlocale(locale.LC_ALL, '')

# Konfiguration und Einstellungen

def load_config(config_path):
    """Lädt die Konfigurationsdatei."""
    config = ConfigParser()
    config.read(config_path)
    return config

# Funktion zur Ausgabe des Fortschritts in einer einzigen Zeile

def print_progress(message: str):
    frame = inspect.currentframe().f_back
    filename = os.path.basename(frame.f_code.co_filename)
    line_number = frame.f_lineno
    function_name = frame.f_code.co_name
    progress_message = f"{filename}:{line_number} [{function_name}] {message}"
    sys.stdout.write(f"\r{progress_message}")
    sys.stdout.flush()

def get_name_from_id(url, headers, endpoint, id):
    response = requests.get(f"{url}/{endpoint}/{id}/", headers=headers)
    if response.status_code == 200:
        return response.json().get("name", "")
    return "Unbekannt"

# Funktion zum Abrufen aller Dokumente

def get_all_documents(url, headers):
    """Holt alle Dokumente von der API ab."""
    documents = []
    page = 1
    while True:
        response = requests.get(
            f"{url}/documents/?page_size=25&page={page}", headers=headers)
        if response.status_code != 200:
            print(
                f"Fehler beim Abrufen der Dokumente. Status Code: {response.status_code}")
            break

        data = response.json()
        documents.extend(data["results"])
        if not data["next"]:
            break
        page += 1
    return documents

# Funktion zum Abrufen der Custom-Field-Definitionen und Mapping erstellen

def get_custom_field_definitions(url, headers):
    """Holt die Definitionen für Custom Fields und erzeugt Mappings für Namen, Typ und Auswahloptionen."""
    print( "Custom Fields")
    custom_fields_response = requests.get(
        f"{url}/custom_fields/", headers=headers)
    custom_fields_map = {}
    custom_field_choices_map = {}

    if custom_fields_response.status_code == 200:
        try:
            custom_fields_data = custom_fields_response.json()
            for field in custom_fields_data["results"]:
                field_id = field["id"]
                field_name = field["name"]
                field_type = field["data_type"]

                # Speichern der Felddefinitionen im custom_fields_map
                custom_fields_map[field_id] = {
                    "name": field_name,
                    "type": field_type,
                    "choices": []
                }

                # Wenn es Auswahloptionen gibt, speichern wir sie
                if field_type == "select":
                    # Wir nehmen hier nur die ersten 5 Optionen als Beispiel,
                    # Du kannst die Anzahl anpassen oder nach Bedarf gestalten
                    custom_field_choices_map[field_id] = {
                        idx: option for idx, option in enumerate(field["extra_data"]["select_options"])
                    }
                    # Füge die Optionen dem custom_fields_map hinzu
                    custom_fields_map[field_id]["choices"] = custom_field_choices_map[field_id]
        except json.decoder.JSONDecodeError as e:
            print_progress(
                f"JSON-Dekodierungsfehler beim Abrufen der Custom Fields: {e}")
            exit()
    else:
        print_progress(
            f"Fehler beim Abrufen der Custom Fields. Status Code: {custom_fields_response.status_code}")
        exit()

    return custom_fields_map

def export_pdf(doc_id, doc_title, tag_directory, url, headers):
    """Lädt das PDF eines Dokuments herunter und speichert es im angegebenen Verzeichnis."""
    pdf_path = os.path.join(tag_directory, f"{doc_title}.pdf")
    pdf_response = requests.get(
        f"{url}/documents/{doc_id}/download/", headers=headers)
    if pdf_response.status_code == 200:
        with open(pdf_path, "wb") as pdf_file:
            pdf_file.write(pdf_response.content)
    else:
        print(f"Fehler beim Herunterladen der PDF für Dokument {doc_title}")

def export_json(doc_data, doc_title, tag_directory):
    """Speichert die Metadaten eines Dokuments als JSON im angegebenen Verzeichnis."""
    json_path = os.path.join(tag_directory, f"{doc_title}.json")
    with open(json_path, "w", encoding="utf-8") as json_file:
        json.dump(doc_data, json_file, ensure_ascii=False, indent=4)
    # print(f"JSON für Dokument {doc_title} gespeichert unter {json_path}")

# Hauptfunktion

def main():
    parser = argparse.ArgumentParser(
        description="Paperless-ngx Dokumentexporter.")
    parser.add_argument("-c", "--config", default="my.ini",
                        help="Pfad zur Konfigurationsdatei")
    args = parser.parse_args()

    # Konfiguration laden
    config = load_config(args.config)
    url = config['paperless']['url']
    token = config['paperless']['token']
    export_directory = config['paperless']['export_directory']
    tag_from_ini = config['paperless']['tags']
    headers = {"Authorization": f"Token {token}"}

    # Custom-Field-Definitionen laden
    print_progress("Lade Custom-Field-Definitionen")
    custom_fields_map = get_custom_field_definitions(
        url, headers)

    # Abrufen der Tags
    print_progress("Lade Tags")
    tags_response = requests.get(f"{url}/tags/", headers=headers)

    # Überprüfen, ob die Anfrage erfolgreich war
    if tags_response.status_code == 200:
       # Tag-Daten extrahieren und in ein Dictionary umwandeln
       tag_dict = {tag["id"]: tag["name"] for tag in tags_response.json()["results"]}

        # Tags direkt ausgeben
       #print("Tags:")
        #for tag_id, tag_name in tag_dict.items():
        #    print(f"ID: {tag_id}, Name: {tag_name}")
    else:
       print(f"Fehler beim Abrufen der Tags: {tags_response.status_code} - {tags_response.text}")

    # Beispiel für einen Tag
    tag_id = next((tid for tid, tname in tag_dict.items()
                  if tname.lower() == tag_from_ini.lower()), None)
    if not tag_id:
        print(f"Tag '{tag_from_ini}' nicht gefunden.")
        exit()

    # Dokumente für Tag abrufen
    documents = get_all_documents(url, headers)
   # Exportieren in Excel
    export_documents_by_tag(tag_from_ini, tag_id, tag_dict, documents, url, headers,
                            custom_fields_map, export_directory)

# Funktion zur Bestimmung der Terminalbreite

def get_terminal_width():
    return os.get_terminal_size().columns

# Exportieren in Excel
def export_documents_by_tag(tag_from_ini, tag_id, tag_dict, documents, url, headers, custom_fields_map, export_directory):
    """Exportiert Dokumente für einen bestimmten Tag in eine Excel-Datei."""
    # Verzeichnis erstellen und alte Dateien löschen
    tag_directory = os.path.join(export_directory, f"export-{tag_from_ini}")
    if os.path.exists(tag_directory):
        for f in os.listdir(tag_directory):
            os.remove(os.path.join(tag_directory, f))
    else:
        os.makedirs(tag_directory)

    # Daten für Excel vorbereiten
# In deinem Code
    terminal_width = get_terminal_width()  # Hole die Breite des Terminals
    document_data = []

# Beispiel für die Verwendung von tqdm
    for doc in tqdm(documents, desc=f"Verarbeite Dokumente für Tag '{tag_from_ini}'", unit="Dok", ncols=terminal_width):
        if tag_id not in doc.get("tags", []):
            continue

        # Custom Fields und weitere Felder aufbereiten
        detailed_doc_response = requests.get(
            f"{url}/documents/{doc['id']}/", headers=headers)
        detailed_doc = detailed_doc_response.json()

        custom_fields = {}
        if "custom_fields" in detailed_doc:
            for custom_field in detailed_doc.get("custom_fields", []):
                field_id = custom_field["field"]
                field_info = custom_fields_map.get(field_id, {})
                field_name = field_info.get(
                    "name", f"Feld {field_id}")  # Hier den Namen abholen
                field_type = field_info.get("type", "string")
                field_value = custom_field["value"]

               # Prüfen, ob der Typ `monetary` ist, um die Währungsformatierung anzuwenden
                if field_type == "monetary":
                    formatted_value = format_currency(field_value)
                    custom_fields[field_name] = formatted_value
                elif field_type == "select":  # Verwendung von `elif` hier
                    # Für Auswahlfelder den Namen des Optionswertes abrufen
                    custom_fields[field_name] = field_info["choices"].get(
                        field_value, f"Wert {field_value}"
                    )
                else:
                    # Standardbehandlung für andere Typen
                    custom_fields[field_name] = field_value

        # Excel-Reihen
        row = {
            "ID": doc.get("id"),
            "Titel": doc.get("title"),
            "Korrespondent": get_name_from_id(url, headers, "correspondents", doc.get("correspondent")),
            "Dokumenttyp": get_name_from_id(url, headers, "document_types", doc.get("document_type")),
            "Speicherpfad": get_name_from_id(url, headers, "storage_paths", doc.get("storage_path")),
            "Tags": ", ".join(tag_dict.get(tag_id, f"Tag {tag_id}") for tag_id in doc.get("tags", [])),
            # Beispiel für das Parsen des Datums mit Zeit und Zeitzoneninformationen
            "Datum": parse_date(doc.get("created")),  # Verwende die parse_date Funktion hier
            "Tags": ", ".join(tag_dict.get(tag_id, f"Tag {tag_id}")
                              for tag_id in doc.get("tags", [])),
            **custom_fields
        }
        # PDF und JSON für jedes Dokument speichern
#        export_pdf(doc['id'], doc['title'], tag_directory, url, headers)
#        export_json(detailed_doc, doc['title'], tag_directory)
        try:
        # Export PDF and JSON
            export_pdf(doc['id'], doc['title'], tag_directory, url, headers)
            export_json(detailed_doc, doc['title'], tag_directory)
        
            # Append document data to rows for Excel or further processing
            document_data.append(row)
        except Exception as e:
            print(f"Failed to export document {doc['id']} ({doc['title']}): {e}")

    # Excel exportieren
    filename = f"export-{tag_from_ini}-{datetime.now().strftime('%Y%m%d')}.xls"
    fullfilename = os.path.join(tag_directory, filename)

    df = pd.DataFrame(document_data)
    with pd.ExcelWriter(fullfilename, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dokumentenliste")
        worksheet = writer.sheets["Dokumentenliste"]
        header_font = Font(bold=True, color="FFFFFF", name="Arial")
        fill = PatternFill(start_color="4F81BD",
                           end_color="4F81BD", fill_type="solid")

     # Kopfzeile formatieren
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = fill

    print(f"\nExcel-Datei erfolgreich erstellt: {fullfilename}")

def parse_date(date_string):
    """
    Versucht, das Datum mit oder ohne Zeitzonen-Offset zu parsen.
    Gibt das Datum im Format '%d.%m.%Y' zurück oder None, wenn das Parsing fehlschlägt.
    """
    if date_string:
        try:
            # Versuche, das Datum mit Zeitzonen-Offset zu parsen
            parsed_date = datetime.strptime(date_string, "%Y-%m-%dT%H:%M:%S%z")
        except ValueError:
            try:
                # Falls das Datum keinen Zeitzonen-Offset hat, versuche es im UTC-Format
                parsed_date = datetime.strptime(date_string, "%Y-%m-%dT%H:%M:%SZ")
            except ValueError:
                # Wenn das Datum in keinem der Formate geparst werden kann, gib None zurück
                parsed_date = None
        
        # Wenn das Datum erfolgreich geparst wurde, formatiere es und gib es zurück
        if parsed_date:
            return parsed_date.strftime("%d.%m.%Y")
    return None


def export_json(document_data, doc_title, tag_directory):
    """
    Export document details as a JSON file.
    """
    # Sanitize the title to avoid invalid file names
    sanitized_title = sanitize_filename(doc_title)
    
    # Build the JSON file path
    json_path = os.path.join(tag_directory, f"{sanitized_title}.json")
    
    # Ensure the directory exists
    os.makedirs(tag_directory, exist_ok=True)
    
    # Write JSON data to file
    try:
        with open(json_path, "w", encoding="utf-8") as json_file:
            json.dump(document_data, json_file, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Error exporting JSON for document {doc_title}: {e}")


def export_pdf(doc_id, doc_title, tag_directory, url, headers):
    """
    Export a PDF for a document.
    """
    # Sanitize the title to avoid issues with special characters
    sanitized_title = sanitize_filename(doc_title)
    
    # Build the PDF path
    pdf_path = os.path.join(tag_directory, f"{sanitized_title}.pdf")
    
    # Ensure the target directory exists
    os.makedirs(tag_directory, exist_ok=True)
    
    # Simulated PDF writing (replace with actual download logic)
    try:
        with open(pdf_path, "wb") as pdf_file:
            # Here would go the logic to write the PDF file from the request
            pass  # Replace with actual PDF content logic
    except Exception as e:
        print(f"Error exporting PDF for document {doc_id}: {e}")


def format_currency(value, currency_locale="de_DE.UTF-8"):
    """Formatiert eine Währungszahl gemäß der angegebenen Locale."""
    if value is None:
        # Log the issue and return an empty string or a default value
        print("Warning: Received None value in format_currency. Returning empty string.")
        return ""
    try:
        # Entferne den Währungsindikator (z.B. EUR) aus dem Wert
        # Nur die Ziffern extrahieren
        clean_value = ''.join(filter(str.isdigit, value))
        if not clean_value:  # Falls keine Ziffern gefunden werden
            return "0,00"  # Standardwert, falls kein gültiger Wert vorhanden ist

        # Wandle den bereinigten Wert in einen float um
        # Angenommen, die Eingabe ist in Cent, daher Division durch 100
        value_float = float(clean_value) / 100
    except ValueError:
        value_float = 0.0  # Standardwert, falls Parsing fehlschlägt

    # Setze die Locale für die Währungsformatierung
    locale.setlocale(locale.LC_ALL, currency_locale)

    # Formatiere den Wert als Währung
    formatted_value = locale.currency(value_float, grouping=True)
    formatted_value = value_float
    return formatted_value

import re

def sanitize_filename(filename):
    """
    Remove or replace characters in the filename that are not allowed in file names.
    """
    sanitized = re.sub(r'[<>:"/\\|?*]', '-', filename)  # Ersetze verbotene Zeichen durch '-'
    return sanitized[:255]  # Truncate to avoid overly long filenames


# Beispielaufruf
#formatted_value = format_currency('EUR22.50')
#print(formatted_value)  # Gibt "22,50 €" aus

# Ausführung der Hauptfunktion
if __name__ == "__main__":
    main()
