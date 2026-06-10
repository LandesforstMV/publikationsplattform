#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel → JSON Konverter für Landesforst MV Publikationen

Verwendung:
    python convert_excel.py input.xlsx output.json
    python convert_excel.py input.csv output.json --csv
    python convert_excel.py "sachgebiet Sitemap.txt" papers.json --csv

Abhängigkeiten:
    pip install pandas openpyxl

Autor: Landesforst MV
"""

import pandas as pd
import json
import sys
import argparse
import re
from pathlib import Path


def clean_text(text):
    """Bereinigt Text von Excel-Artefakten."""
    if pd.isna(text) or text is None:
        return ""
    text = str(text).strip()
    # Entferne typische Excel-NaN-Werte
    if text.lower() in ['nan', 'nat', 'none', '']:
        return ""
    return text


def parse_authors(authors_str):
    """Parst Autorenstring in Liste."""
    if not authors_str:
        return []

    # Ersetze verschiedene Trennzeichen durch Komma
    separators = [';', ' & ', ' und ', ' and ', ' / ']
    for sep in separators:
        authors_str = authors_str.replace(sep, ', ')

    # Splitte und bereinige
    authors = []
    for author in authors_str.split(','):
        author = author.strip()
        if author and len(author) > 1:
            authors.append(author)

    return authors


def extract_year(year_value):
    """Extrahiert Jahr aus verschiedenen Formaten."""
    if pd.isna(year_value):
        return 0

    # Wenn es bereits eine Zahl ist
    if isinstance(year_value, (int, float)):
        return int(year_value)

    # Aus String extrahieren
    year_str = str(year_value).strip()
    match = re.search(r'(19|20)\d{2}', year_str)
    if match:
        return int(match.group())

    return 0


def is_valid_link(link):
    """Prüft, ob ein Link gültig ist."""
    if not link:
        return False

    invalid_patterns = [
        'nan', 'NaN', 'pdf liegt nicht vor', 'nicht öffentlich',
        'darf nicht veröffentlicht werden', 'bitte an frau abraham wenden',
        'unveröffentl.', 'unveröff.', 'unveröffent.', 'unveröff.'
    ]

    link_lower = link.lower()
    for pattern in invalid_patterns:
        if pattern in link_lower:
            return False

    return True


def convert_excel_to_json(input_file, output_file):
    """Hauptkonvertierungsfunktion."""

    print(f"📖 Lese {input_file}...")

    # Datei einlesen
    file_path = Path(input_file)
    if file_path.suffix.lower() in ['.csv', '.txt']:
        df = pd.read_csv(input_file, sep='\t', encoding='utf-8')
    else:
        df = pd.read_excel(input_file, engine='openpyxl')

    print(f"✅ {len(df)} Zeilen gelesen")

    # Spalten bereinigen
    df.columns = df.columns.str.strip()

    # Erwartete Spalten (flexibel)
    column_mapping = {
        'sachgebiet': ['sachgebiet', 'Sachgebiet', 'SACHGEBIET', 'Fachgebiet'],
        'sitemap': ['Sitemap', 'sitemap', 'SITEMAP', 'Kategorie', 'kategorie'],
        'projekt': ['Projekt', 'projekt', 'PROJEKT'],
        'dok_typ': ['dok_typ', 'Dokumenttyp', 'dokumenttyp', 'Typ', 'typ'],
        'autor': ['autor', 'Autor', 'AUTOR', 'autoren', 'Autoren'],
        'jahr': ['jahr', 'Jahr', 'JAHR', 'year', 'Year'],
        'titel': ['titel', 'Titel', 'TITEL', 'title', 'Title'],
        'veroeffentlicht': ['veröffentlicht', 'veroeffentlicht', 'Veroeffentlicht', 
                           'Publikation', 'publikation', 'Journal', 'journal'],
        'link': ['Link', 'link', 'LINK', 'URL', 'url', 'Pdf', 'pdf']
    }

    # Finde tatsächliche Spaltennamen
    actual_columns = {}
    for standard, alternatives in column_mapping.items():
        for col in df.columns:
            if col in alternatives:
                actual_columns[standard] = col
                break

    print(f"🔍 Gefundene Spalten: {list(actual_columns.keys())}")

    # Daten transformieren
    papers = []

    for idx, row in df.iterrows():
        # Titel ist Pflichtfeld
        titel = clean_text(row.get(actual_columns.get('titel', 'titel'), ''))
        if not titel:
            continue

        # Jahr extrahieren
        jahr_col = actual_columns.get('jahr', 'jahr')
        jahr = extract_year(row.get(jahr_col, 0))
        if jahr == 0:
            continue

        # Autoren parsen
        autor_col = actual_columns.get('autor', 'autor')
        autoren = parse_authors(clean_text(row.get(autor_col, '')))

        # Link prüfen
        link_col = actual_columns.get('link', 'link')
        link = clean_text(row.get(link_col, ''))
        has_valid_link = is_valid_link(link)

        paper = {
            'id': f"paper-{idx}",
            'titel': titel,
            'autoren': autoren,
            'jahr': jahr,
            'sachgebiet': clean_text(row.get(actual_columns.get('sachgebiet', 'sachgebiet'), '')),
            'sitemap': clean_text(row.get(actual_columns.get('sitemap', 'sitemap'), '')),
            'projekt': clean_text(row.get(actual_columns.get('projekt', 'projekt'), '')),
            'dok_typ': clean_text(row.get(actual_columns.get('dok_typ', 'dok_typ'), 'Publikation')),
            'veroeffentlicht': clean_text(row.get(actual_columns.get('veroeffentlicht', 'veroeffentlicht'), '')),
            'link': link if has_valid_link else '',
            'aktiv': 'ja'
        }

        papers.append(paper)

    # Sortieren: Neueste zuerst
    papers.sort(key=lambda x: x['jahr'], reverse=True)

    # JSON speichern
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(papers, f, ensure_ascii=False, indent=2)

    # Statistik
    years = set(p['jahr'] for p in papers)
    sachgebiete = set(p['sachgebiet'] for p in papers if p['sachgebiet'])
    typs = set(p['dok_typ'] for p in papers if p['dok_typ'])
    with_links = sum(1 for p in papers if p['link'])

    print(f"\n✅ Konvertierung erfolgreich!")
    print(f"   📚 {len(papers)} Publikationen")
    print(f"   📅 {len(years)} Jahre ({min(years)}–{max(years)})")
    print(f"   📂 {len(sachgebiete)} Sachgebiete")
    print(f"   📄 {len(typs)} Dokumenttypen")
    print(f"   🔗 {with_links} mit PDF-Link")
    print(f"\n💾 Gespeichert als: {output_file}")

    return papers


def main():
    parser = argparse.ArgumentParser(
        description='Konvertiert Excel/CSV Publikationsliste zu JSON'
    )
    parser.add_argument('input', help='Eingabedatei (.xlsx, .csv, .txt)')
    parser.add_argument('output', nargs='?', default='papers.json', 
                       help='Ausgabedatei (default: papers.json)')
    parser.add_argument('--csv', action='store_true', 
                       help='Eingabe ist CSV/TSV (tab-getrennt)')

    args = parser.parse_args()

    try:
        convert_excel_to_json(args.input, args.output)
    except Exception as e:
        print(f"❌ Fehler: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
