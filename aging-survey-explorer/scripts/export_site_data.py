"""Export one JSON row per respondent for the website (data/survey_responses.json).
Run after build.py and add_tags.py (reads combined.pkl and tags.pkl from the current folder)."""
import json, math, sys
import pandas as pd

OUT = sys.argv[1] if len(sys.argv) > 1 else 'data/survey_responses.json'
df = pd.read_pickle('combined.pkl')
T = pd.read_pickle('tags.pkl')
tags, qs = T['tags'], T['questions']
US = 'United States'

def region(c):
    if not isinstance(c, str): return None
    return US if c == US else ('China' if c == 'China' else 'Other countries')

def agecat(a, student_max=29, elderly_min=65):
    if a is None: return None
    if a >= elderly_min: return f'Elderly ({elderly_min}+)'
    if a <= student_max: return f'Student age (≤{student_max})'
    return f'Middle-aged ({student_max + 1}–{elderly_min - 1})'

def gender(g):
    if not isinstance(g, str): return None
    return g if g in ('Female', 'Male') else 'Another / not stated'

def mobility(f, r):
    if not isinstance(r, str): return None
    if r != US: return 'Lives outside US'
    if f == US: return 'From US, lives in US'
    return 'From abroad, lives in US' if isinstance(f, str) else None

def v(x):
    if x is None: return None
    if isinstance(x, float):
        if math.isnan(x): return None
        return int(x) if x.is_integer() else round(x, 2)
    return x

rows = []
for i, r in df.reset_index(drop=True).iterrows():
    age = v(r['Age'])
    o = {'id': r['Respondent ID'], 'group': 'Group %d' % r['Group'], 'age': age, 'age_cat': agecat(age),
         'gender': gender(r['Gender']), 'res_region': region(r['Residence – Country']),
         'from_region': region(r['From – Country']), 'mobility': mobility(r['From – Country'], r['Residence – Country']),
         'res_place': v(r['Residence (State, Country)']), 'from_place': v(r['From (State, Country)']),
         'pct_est': v(r['Q6 Est. % of US population 65+']), 'pct_bin': v(r['Q6 Estimate (range)']),
         'rank_p1': v(r['Personal rank – Reproductive health & longevity']),
         'rank_p2': v(r['Personal rank – Democratizing longevity']),
         'rank_p3': v(r['Personal rank – AI/ML for aging & health']),
         'rank_s1': v(r['Societal rank – Reproductive health & longevity']),
         'rank_s2': v(r['Societal rank – Democratizing longevity']),
         'rank_s3': v(r['Societal rank – AI/ML for aging & health']),
         'excite_topic': v(r['Most exciting topic']), 'concern_topic': v(r['Most concerning topic']),
         'ethics_yn': v(r['Ethical considerations needed?']), 'excited_yn': v(r['Excited about future of aging research?']),
         'excited_r': v(r['Excitement rating (1–5)']), 'bio_yn': v(r['Bioengineers can help?']),
         'bio_r': v(r['Bioengineer role rating (1–5)'])}
    for k, col, _ in qs:
        o['t_' + k] = '|'.join(tags[k].iloc[i]) or None
        t = r[col]
        o['x_' + k] = t[:600] if isinstance(t, str) else None
    o['res_country'] = v(r['Residence – Country']); o['res_state'] = v(r['Residence – State/Province'])
    o['from_country'] = v(r['From – Country']); o['from_state'] = v(r['From – State/Province'])
    rows.append(o)
json.dump(rows, open(OUT, 'w'), ensure_ascii=False)
print(len(rows), 'rows ->', OUT)
