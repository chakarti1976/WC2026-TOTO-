#!/usr/bin/env python3
"""
Extracts data from the WC 2026 Toto Excel file and generates data.js for the web app.
Run this script whenever you update the Excel file, or set up GitHub Actions to run it automatically.
"""
import pandas as pd
import json
import os
import sys

EXCEL_FILE = 'WC 2026 Toto - Participants list.xlsx'

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

PICK_MULTIPLIERS = [
    ('free1', 6),
    ('free2', 4),
    ('free3', 3),
    ('free4', 2),
    ('oc1',   1),
    ('oc2',   1),
    ('oc3',   1),
]

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
    stats = {}
    for group, teams in ALL_GROUPS.items():
        for t in teams:
            stats[t] = {'team': t, 'group': group, 'pl': 0, 'w': 0, 'd': 0, 'l': 0, 'gf': 0, 'ga': 0, 'pts': 0}

    for m in matches:
        if not m['played']:
            continue
        t1, t2 = m['team1'], m['team2']
        s1, s2 = int(m['score1']), int(m['score2'])
        if t1 not in stats or t2 not in stats:
            continue
        stats[t1]['pl'] += 1; stats[t2]['pl'] += 1
        stats[t1]['gf'] += s1; stats[t1]['ga'] += s2
        stats[t2]['gf'] += s2; stats[t2]['ga'] += s1
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
            'team': s['team'], 'pl': s['pl'], 'w': s['w'], 'd': s['d'], 'l': s['l'],
            'gd': f"{s['gf']} - {s['ga']}", 'pts': s['pts']
        } for s in group_teams]

    return result, stats

def compute_team_goals(matches):
    """Total goals scored by each team across all played matches."""
    goals = {}
    for m in matches:
        if not m['played']:
            continue
        t1, t2 = m['team1'], m['team2']
        goals[t1] = goals.get(t1, 0) + int(m['score1'])
        goals[t2] = goals.get(t2, 0) + int(m['score2'])
    return goals

def compute_team_results(matches):
    """
    Returns a dict: team -> {'wins': n, 'draws': n} across all played matches.
    Used to calculate group stage points per team pick.
    """
    results = {}
    for m in matches:
        if not m['played']:
            continue
        t1, t2 = m['team1'], m['team2']
        s1, s2 = int(m['score1']), int(m['score2'])
        for t in [t1, t2]:
            if t not in results:
                results[t] = {'wins': 0, 'draws': 0}
        if s1 > s2:
            results[t1]['wins'] += 1
        elif s2 > s1:
            results[t2]['wins'] += 1
        else:
            results[t1]['draws'] += 1
            results[t2]['draws'] += 1
    return results

def compute_leaderboard(toto_entries, matches):
    """
    Calculate points for each participant based on:
    - Free picks (×6, ×4, ×3, ×2): 3pts per win, 1pt per draw, times multiplier
    - OC picks (×1): same scoring
    - Best scoring team (×0.5 per goal scored by that team)
    Returns updated toto_entries with computed totals and ranks.
    """
    team_results = compute_team_results(matches)
    team_goals = compute_team_goals(matches)

    for p in toto_entries:
        total = 0.0
        breakdown = {}

        for field, multiplier in PICK_MULTIPLIERS:
            team = p.get(field, '')
            if not team:
                continue
            tr = team_results.get(team, {'wins': 0, 'draws': 0})
            pts = (tr['wins'] * 3 + tr['draws'] * 1) * multiplier
            breakdown[field] = round(pts, 2)
            total += pts

        # Best scoring team bonus
        scoring_team = p.get('scoring_team', '')
        if scoring_team:
            goals = team_goals.get(scoring_team, 0)
            scoring_pts = goals * 0.5
            breakdown['scoring'] = round(scoring_pts, 2)
            total += scoring_pts

        p['total'] = round(total, 2)
        p['breakdown'] = breakdown

    # Rank by total descending
    sorted_entries = sorted(toto_entries, key=lambda x: (-x['total'], x['name']))
    rank = 1
    for i, p in enumerate(sorted_entries):
        if i > 0 and p['total'] == sorted_entries[i-1]['total']:
            p['rank'] = sorted_entries[i-1]['rank']
        else:
            p['rank'] = rank
        rank = i + 2

    return sorted_entries

def extract_toto_teams(df_t, matches):
    toto = []
    skip_names = {'Participant','FIFA WC 2026 — Toto Participants','-','Multipliers →',''}
    for idx, row in df_t.iterrows():
        name = row[0]
        if pd.isna(name): continue
        name_s = str(name).strip()
        if name_s in skip_names: continue
        free1 = str(row[1]).strip() if pd.notna(row[1]) else ''
        if not free1 or free1 in ['Free 1 (×6)','nan','-','×6']: continue
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
            'total': 0,
            'rank': 0,
        })

    return compute_leaderboard(toto, matches)

def build_results_leaderboard(toto):
    """Convert computed toto scores into the results/leaderboard format."""
    results = []
    for p in toto:
        bd = p.get('breakdown', {})
        group_pts = round(
            bd.get('free1', 0) + bd.get('free2', 0) +
            bd.get('free3', 0) + bd.get('free4', 0) +
            bd.get('oc1', 0) + bd.get('oc2', 0) + bd.get('oc3', 0), 2
        )
        results.append({
            'rank': p['rank'],
            'name': p['name'],
            'group': group_pts,
            'r32': 0,
            'r16': 0,
            'qf': 0,
            'sf': 0,
            'final_pts': 0,
            'third': 0,
            'champ': 0,
            'total': p['total'],
            'prize': 0,
        })
    results.sort(key=lambda x: (x['rank'], x['name']))
    return results

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"ERROR: {EXCEL_FILE} not found.")
        sys.exit(1)

    print(f"Reading {EXCEL_FILE}...")

    df_main = pd.read_excel(EXCEL_FILE, sheet_name='FIFA WC 2026', header=None)
    df_toto = pd.read_excel(EXCEL_FILE, sheet_name=' Toto teams', header=None)

    matches = extract_matches(df_main)
    groups, _ = compute_group_standings(matches)
    toto = extract_toto_teams(df_toto, matches)
    results = build_results_leaderboard(toto)

    played = sum(1 for m in matches if m['played'])
    top5 = toto[:5]

    data = {
        'matches': matches,
        'groups': groups,
        'results': results,
        'toto': toto,
        'last_updated': pd.Timestamp.now().strftime('%Y-%m-%d %H:%M UTC'),
        'total_pot': 3150,
        'tournament_name': 'FIFA World Cup 2026'
    }

    print(f"  Matches: {len(matches)} ({played} played)")
    print(f"  Groups: {len(groups)}")
    print(f"  Toto entries: {len(toto)}")
    print(f"  Top 5 leaderboard:")
    for p in top5:
        print(f"    #{p['rank']} {p['name']}: {p['total']} pts")

    js_content = f"// Auto-generated by extract_data.py — do not edit manually\nconst WC_DATA = {json.dumps(data, ensure_ascii=False, indent=2)};\n"

    with open('data.js', 'w', encoding='utf-8') as f:
        f.write(js_content)

    print("✓ Generated data.js successfully")

if __name__ == '__main__':
    main()
