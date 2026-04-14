#!/usr/bin/env python3
"""
ECS Drug Checking — Script de mise à jour du dashboard
=======================================================
Usage : double-cliquer sur ce fichier (ou python update_dashboard.py)

Ce script :
1. Lit le fichier Excel local (Re_sultats_ECS.xlsx)
2. Calcule toutes les statistiques
3. Met à jour le fichier index.html avec les nouvelles données
4. Pousse automatiquement sur GitHub Pages

Prérequis (à installer une seule fois) :
  pip install pandas openpyxl gitpython
"""

import pandas as pd
import numpy as np
import json
import re
import os
import sys
from datetime import datetime
from pathlib import Path

# ============================================================
# CONFIGURATION — à adapter à ton environnement
# ============================================================

# Chemin vers ton fichier Excel (absolu ou relatif à ce script)
EXCEL_PATH = "Re_sultats_ECS.xlsx"

# Chemin vers le dossier GitHub cloné localement
# ex: "C:/Users/Pierre/Documents/ESC_Monitoring"
GITHUB_REPO_PATH = "."  # par défaut : même dossier que ce script

# Nom du fichier HTML dans le repo GitHub
HTML_FILENAME = "index.html"

# ============================================================
# FONCTIONS
# ============================================================

def normalize_product(p):
    p = str(p).lower()
    if 'héroïne' in p or 'heroine' in p or 'héro' in p: return 'Héroïne'
    if 'base' in p and 'coca' in p: return 'Cocaine base'
    if 'crack' in p: return 'Cocaine base'
    if 'hcl' in p or ('coca' in p and 'base' not in p): return 'Cocaine HCl'
    return 'Autre'

def aggregate_by_period(df, date_col, val_col, period):
    df = df.copy()
    df['val'] = pd.to_numeric(df[val_col], errors='coerce')
    df = df.dropna(subset=['val', date_col])
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df = df.dropna(subset=[date_col])

    if period == 'mois':
        df['key'] = df[date_col].dt.to_period('M').astype(str)
    elif period == 'trimestre':
        df['key'] = df[date_col].dt.to_period('Q').astype(str)
    else:
        df['key'] = df[date_col].dt.year.astype(str)

    result = []
    for key, group in df.groupby('key'):
        vals = group['val'].dropna().sort_values()
        n = len(vals)
        if n == 0: continue
        mean = float(vals.mean())
        median = float(vals.median())
        q25 = float(vals.quantile(0.25))
        q75 = float(vals.quantile(0.75))
        std = float(vals.std()) if n > 1 else None
        row = {period: key, 'count': n, 'mean': round(mean,2),
               'median': round(median,2), 'q25': round(q25,2),
               'q75': round(q75,2), 'std': round(std,2) if std else None}
        result.append(row)

    return sorted(result, key=lambda x: x[period])

def get_purity_data(df, date_col, val_col):
    return {
        'mois': aggregate_by_period(df, date_col, val_col, 'mois'),
        'trimestre': aggregate_by_period(df, date_col, val_col, 'trimestre'),
        'annee': aggregate_by_period(df, date_col, val_col, 'annee'),
    }

def compute_data(excel_path):
    print(f"  → Lecture de {excel_path}...")

    df_coke = pd.read_excel(excel_path, sheet_name='cocaine_HCl', header=0)
    df_base = pd.read_excel(excel_path, sheet_name='cocaine_base', header=0)
    df_hero = pd.read_excel(excel_path, sheet_name='Hero', header=0)

    # Filtrer les années futures aberrantes
    for df in [df_coke, df_base, df_hero]:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    df_coke = df_coke[df_coke['Date'].dt.year <= datetime.now().year + 1]
    df_base = df_base[df_base['Date'].dt.year <= datetime.now().year + 1]
    df_hero = df_hero[df_hero['Date'].dt.year <= datetime.now().year + 1]

    purity_coke_hcl = get_purity_data(df_coke, 'Date', 'pureté HCl')
    purity_coke_base = get_purity_data(df_base, 'Date', 'pureté Base')
    purity_hero = get_purity_data(df_hero, 'Date', 'pureté Base')

    # Échantillons par mois
    counts = {}
    for df, cat in [(df_coke,'Cocaine HCl'), (df_base,'Cocaine base'), (df_hero,'Héroïne')]:
        df2 = df.dropna(subset=['Date'])
        for _, row in df2.iterrows():
            mois = row['Date'].strftime('%Y-%m')
            key = f"{mois}||{cat}"
            counts[key] = counts.get(key, 0) + 1

    samples_by_month = sorted([
        {'mois': k.split('||')[0], 'cat': k.split('||')[1], 'count': v}
        for k, v in counts.items()
    ], key=lambda x: x['mois'])

    totals = {
        'coke_hcl': len(df_coke.dropna(subset=['Date'])),
        'coke_base': len(df_base.dropna(subset=['Date'])),
        'hero': len(df_hero.dropna(subset=['Date'])),
    }
    totals['total_samples'] = sum(totals.values())

    return {
        'purity_coke_hcl': purity_coke_hcl,
        'purity_coke_base': purity_coke_base,
        'purity_hero': purity_hero,
        'samples_by_month': samples_by_month,
        'totals': totals,
        'last_updated': datetime.now().isoformat()
    }

def update_html(html_path, data):
    print(f"  → Mise à jour de {html_path}...")
    with open(html_path, 'r', encoding='utf-8') as f:
        content = f.read()

    json_str = json.dumps(data, ensure_ascii=False, default=str)

    # Remplace le bloc RAW_STATIC
    pattern = r'(const RAW_STATIC = )\{.*?\};'
    replacement = f'const RAW_STATIC = {json_str};'
    new_content = re.sub(pattern, replacement, content, flags=re.DOTALL)

    if new_content == content:
        print("  ⚠️  Pattern RAW_STATIC non trouvé dans le HTML — vérifier le fichier")
        return False

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    print(f"  ✓ HTML mis à jour ({totals_str(data)})")
    return True

def totals_str(data):
    t = data['totals']
    return f"n={t['total_samples']} (HCl:{t['coke_hcl']} Base:{t['coke_base']} Héro:{t['hero']})"

def push_to_github(repo_path):
    try:
        from git import Repo
        print(f"  → Push vers GitHub...")
        repo = Repo(repo_path)
        repo.git.add(all=True)
        msg = f"Dashboard mis à jour — {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        repo.index.commit(msg)
        origin = repo.remote(name='origin')
        origin.push()
        print(f"  ✓ Publié sur GitHub Pages")
        return True
    except ImportError:
        print("  ⚠️  Module 'git' non installé. Installe-le avec: pip install gitpython")
        print("  → Le fichier HTML a été mis à jour localement.")
        print("  → Upload-le manuellement sur GitHub pour publier.")
        return False
    except Exception as e:
        print(f"  ⚠️  Erreur Git: {e}")
        print("  → Le fichier HTML a été mis à jour localement.")
        return False

# ============================================================
# MAIN
# ============================================================

def main():
    print("\n" + "="*50)
    print("  ECS Drug Checking — Mise à jour dashboard")
    print("="*50 + "\n")

    # 1. Trouver le fichier Excel
    script_dir = Path(__file__).parent
    excel_path = script_dir / EXCEL_PATH
    if not excel_path.exists():
        # Chercher dans le dossier courant
        excel_path = Path(EXCEL_PATH)
    if not excel_path.exists():
        print(f"❌ Fichier Excel introuvable : {EXCEL_PATH}")
        print(f"   Place ce script dans le même dossier que {EXCEL_PATH}")
        input("\nAppuie sur Entrée pour fermer...")
        sys.exit(1)

    # 2. Trouver le fichier HTML
    repo_path = script_dir / GITHUB_REPO_PATH
    html_path = repo_path / HTML_FILENAME
    if not html_path.exists():
        html_path = script_dir / HTML_FILENAME
    if not html_path.exists():
        print(f"❌ Fichier HTML introuvable : {HTML_FILENAME}")
        print(f"   Assure-toi que {HTML_FILENAME} est dans le même dossier que ce script")
        input("\nAppuie sur Entrée pour fermer...")
        sys.exit(1)

    try:
        # 3. Calculer les données
        data = compute_data(str(excel_path))

        # 4. Mettre à jour le HTML
        ok = update_html(str(html_path), data)
        if not ok:
            input("\nAppuie sur Entrée pour fermer...")
            sys.exit(1)

        # 5. Pousser sur GitHub
        push_to_github(str(repo_path))

        print(f"\n✅ Dashboard mis à jour avec succès !")
        print(f"   {totals_str(data)}")
        print(f"   🌐 https://pitesseiva-spec.github.io/ESC_Monitoring/index.html")

    except Exception as e:
        print(f"\n❌ Erreur : {e}")
        import traceback
        traceback.print_exc()

    print()
    input("Appuie sur Entrée pour fermer...")

if __name__ == '__main__':
    main()
