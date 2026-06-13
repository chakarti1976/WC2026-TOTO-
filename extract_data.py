#!/usr/bin/env python3
"""
Extracts data from the WC 2026 Toto Excel file and generates data.js for the web app.
Run this script whenever you update the Excel file, or set up GitHub Actions to run it automatically.
"""
import pandas as pd
import json
import os
import sys

EXCEL_FILE = 'WC_2026_Toto__Participants_list.xlsx'

ALL_GROUPS = {
    'A': ['Mexico','South Korea','Czechia','South Africa'],
    'B': ['Canada','Bosnia & Herzegovina','Qatar','Switzerland'],
    'C': ['Brazil','Morocco','Haiti','Scotland'],
    'D': ['United States','Australia','Türkiye','Paraguay'],
    'E': ['Germany','Curaçao','Ivory Coast','Ecuador'],
    'F': ['Netherlands','Japan','Sweden','Tunisia'],
    'G': ['Belgium','Egypt','Iran','New Zealand'],
    'H': ['Spain','Cape Verde','Saudi Arabia','Uruguay'],
    'I': ['France','Senegal','Iraq','Norway'],
    'J': ['Argentina','Algeria','Austria','Jordan'],
    'K': ['Portugal','DR Congo','Uzbekistan','Colombia'],
    'L': ['England','Croatia','Ghana','Panama'],
}

def clean_float(val, default=0):
    try:
        if pd.isna(val): return default
        return float(val)
    except: return default

def extract_matches(df):
    matches = []
    for idx, row in df.iterrows():
        val = row[0]
        if pd.notna(val) and str(val).strip().replace('.0','').isdigit():
            match_num = int(float(val))
            if match_num < 1 or match_num > 72: continue
            score1_raw = row[5]
            score2_raw = row[6]
            s1 = '' if pd.isna(score1_raw) else str(int(float(score1_raw)))
            s2 = '' if pd.isna(score2_raw) else str(int(float(score2_raw)))
            matches.append({
                'match': match_num,
                'day': str(row[1]) if pd.notna(row[1]) else '',
                'date': str(row[2]) if pd.notna(row[2]) else '',
                'time': str(row[3]) if pd.notna(row[3]) else '',
                'team1': str(row[4]).strip() if pd.notna(row[4]) else '',
                'score1': s1,
                'score2': s2,
                'team2': str(row[7]).strip() if pd.notna(row[7]) else '',
                'group': str(row[8]).strip() if pd.notna(row[8]) else '',
                'played': s1 != '' and s2 != ''
            })
    return matches

def compute_group_standings(matches):
    """Compute group standings from match results."""
    team_to_group = {t: g for g, teams in ALL_GROUPS.items() for t in teams}

    # Initialize stats for every team
    stats = {}
    for group, teams in ALL_GROUPS.items():
        for t in teams:
            stats[t] = {'team': t, 'group': group, 'pl': 0, 'w': 0, 'd': 0, 'l': 0, 'gf': 0, 'ga': 0, 'pts': 0}

    for m in matches:
        if not m['played']:
            continue
        t1, t2 = m['team1'], m['team2']
        s1, s2 = int(m['score1']), int(m['score2'])
        for t in [t1, t2]:
            if t not in stats:
                continue
        stats[t1]['pl'] += 1
        stats[t2]['pl'] += 1
        stats[t1]['gf'] += s1
        stats[t1]['ga'] += s2
        stats[t2]['gf'] += s2
        stats[t2]['ga'] += s1
        if s1 > s2:
            stats[t1]['w'] += 1; stats[t1]['pts'] += 3
            stats[t2]['l'] += 1
        elif s2 > s1:
            stats[t2]['w'] += 1; stats[t2]['pts'] += 3
            stats[t1]['l'] += 1
        else:
            stats[t1]['d'] += 1; stats[t1]['pts'] += 1
            stats[t2]['d'] += 1; stats[t2]['pts'] += 1

    result = {}
    for group, teams in ALL_GROUPS.items():
        group_teams = [stats[t] for t in teams]
        group_teams.sort(key=lambda x: (-x['pts'], -(x['gf'] - x['ga']), -x['gf']))
        result[group] = [{
            'team': s['team'],
            'pl': s['pl'],
            'w': s['w'],
            'd': s['d'],
            'l': s['l'],
            'gd': f"{s['gf']} - {s['ga']}",
            'pts': s['pts']
        } for s in group_teams]

    return result

def extract_results(df_res):
    participants = []
    for idx, row in df_res.iterrows():
        rank = row[0]
        name = row[1]
        if pd.notna(rank) and pd.notna(name) and str(rank).strip().replace('.0','').isdigit():
            r = int(float(rank))
            if r > 200: continue
            participants.append({
                'rank': r,
                'name': str(name).strip(),
                'group': clean_float(row[2]),
                'r32': clean_float(row[3]),
                'r16': clean_float(row[4]),
                'qf': clean_float(row[5]),
                'sf': clean_float(row[6]),
                'final_pts': clean_float(row[7]),
                'third': clean_float(row[8]),
                'champ': clean_float(row[9]),
                'total': clean_float(row[10]),
                'prize': round(clean_float(row[11]), 2),
            })
    return participants

def extract_toto_teams(df_t):
    toto = []
    skip_names = {'Participant','FIFA WC 2026 — Toto Participants','-','Multipliers →',''}
    for idx, row in df_t.iterrows():
        name = row[0]
        if pd.isna(name): continue
        name_s = str(name).strip()
        if name_s in skip_names: continue
        free1 = str(row[1]).strip() if pd.notna(row[1]) else ''
        if not free1 or free1 in ['Free 1 (×6)','nan','-','×6']: continue
        total_raw = row[17]
        rank_raw = row[18]
        toto.append({
            'name': name_s,
            'free1': free1,
            'free2': str(row[2]).strip() if pd.notna(row[2]) else '',
            'free3': str(row[3]).strip() if pd.notna(row[3]) else '',
            'free4': str(row[4]).strip() if pd.notna(row[4]) else '',
            'oc1': str(row[5]).strip() if pd.notna(row[5]) else '',
            'oc2': str(row[6]).strip() if pd.notna(row[6]) else '',
            'oc3': str(row[7]).strip() if pd.notna(row[7]) else '',
            'scoring_team': str(row[8]).strip() if pd.notna(row[8]) else '',
            'total': clean_float(total_raw),
            'rank': int(clean_float(rank_raw)) if pd.notna(rank_raw) else 0,
        })
    return toto

def extract_rules(df_rules):
    lines = []
    for idx, row in df_rules.iterrows():
        for col in range(min(10, len(row))):
            val = row[col]
            if pd.notna(val) and str(val).strip():
                lines.append(str(val).strip())
        if idx > 80: break
    return lines

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: {EXCEL_FILE} not found. Place it in the same directory as this script.")
        sys.exit(1)

    print(f"Reading {EXCEL_FILE}...")

    df_main = pd.read_excel(EXCEL_FILE, sheet_name='FIFA WC 2026', header=None)
    df_res = pd.read_excel(EXCEL_FILE, sheet_name='Results', header=None)
    df_toto = pd.read_excel(EXCEL_FILE, sheet_name=' Toto teams', header=None)

    matches = extract_matches(df_main)
    groups = compute_group_standings(matches)

    data = {
        'matches': matches,
        'groups': groups,
        'results': extract_results(df_res),
        'toto': extract_toto_teams(df_toto),
        'last_updated': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M UTC'),
        'total_pot': 3150,
        'tournament_name': 'FIFA World Cup 2026'
    }

    print(f"  Matches: {len(data['matches'])} ({sum(1 for m in matches if m['played'])} played)")
    print(f"  Groups: {len(data['groups'])}")
    print(f"  Results: {len(data['results'])}")
    print(f"  Toto entries: {len(data['toto'])}")

    js_content = f"// Auto-generated by extract_data.py — do not edit manually\nconst WC_DATA = {json.dumps(data, ensure_ascii=False, indent=2)};\n"

    with open('data.js', 'w', encoding='utf-8') as f:
        f.write(js_content)

    print("✓ Generated data.js successfully")

if __name__ == '__main__':
    main()
