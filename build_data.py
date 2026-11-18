#!/usr/bin/env python3
"""Build data.json for the KPN Cash Till Manned Dashboard.

Usage:  python build_data.py [data-folder] [out.json]
Default: every .xlsx/.csv in ./data  ->  ./data.json

- Reads the daily D-1 file ("Cash Till - Store wise D-1" layout: date row on top, header with outlet_id)
  and the monthly day-wise layout (with a Date column). Files are merged by date; when two files carry
  the same date, the newest file wins ("newest" = last git commit time, falls back to file modified time).
- Cluster -> TM mapping: config/tm_master.csv (Cluster Name,TM Name).
- Data checks run on the newest file. Problems are written to data_issues.txt and printed as GitHub
  warnings; the workflow then fails so GitHub emails the repo owner. Critical problems stop the build
  (data.json is left unchanged).
"""
import sys, os, re, csv, glob, json, subprocess, statistics, datetime as dt
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'data')
out = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, 'data.json')
issues_path = os.path.join(HERE, 'data_issues.txt')
SLOTS = [((390 + 30 * i) // 60, (390 + 30 * i) % 60) for i in range(31)]   # 06:30 .. 21:30


def stamp(f):
    try:
        t = subprocess.run(['git', 'log', '-1', '--format=%ct', '--', f], capture_output=True, text=True, cwd=HERE).stdout.strip()
        if t: return (int(t), os.path.getmtime(f))
    except Exception: pass
    return (0, os.path.getmtime(f))


def conv(v):
    """CSV cell -> time / datetime / number / str (xlsx cells already come typed)."""
    if not isinstance(v, str): return v
    v = v.strip()
    m = re.fullmatch(r'(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?', v)
    if m:
        h = int(m.group(1))
        if m.group(3): h = h % 12 + (12 if m.group(3).lower() == 'pm' else 0)
        return dt.time(h, int(m.group(2)))
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y', '%d-%b-%Y', '%d %b %Y', '%Y-%m-%d %H:%M:%S'):
        try: return dt.datetime.strptime(v, fmt)
        except ValueError: pass
    try: return float(v) if '.' in v else int(v)
    except ValueError: return v or None


def read_rows(f):
    if f.lower().endswith('.csv'):
        return [tuple(conv(x) for x in r) for r in csv.reader(open(f, encoding='utf-8-sig'))]
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    ws = wb['Cash Till - Store wise D-1'] if 'Cash Till - Store wise D-1' in wb.sheetnames else wb.worksheets[0]
    return [tuple(conv(x) for x in r) for r in ws.iter_rows(values_only=True)]


class Bad(Exception): pass


def parse(f):
    """-> (records {(date, sid): (planned, slots_str, name, cluster)}, header info)"""
    rows = read_rows(f)
    hi = next((i for i, r in enumerate(rows) if r and 'outlet_id' in [str(x).strip() if x is not None else '' for x in r]), None)
    if hi is None: raise Bad('header row with "outlet_id" not found — layout changed?')
    hdr = rows[hi]
    col = {str(x).strip(): j for j, x in enumerate(hdr) if isinstance(x, str)}
    for need in ('outlet_id', 'Store Name', 'Cluster Name', 'Planned'):
        if need not in col: raise Bad(f'column "{need}" missing')
    tcol = {(x.hour, x.minute): j for j, x in enumerate(hdr) if isinstance(x, dt.time)}
    miss = [f'{h:02d}:{m:02d}' for h, m in SLOTS if (h, m) not in tcol]
    if miss: raise Bad('missing half-hour columns ' + ', '.join(miss))
    fdate = None
    if 'Date' not in col:
        for r in rows[:hi]:
            for x in r or ():
                if isinstance(x, dt.datetime): fdate = x.strftime('%Y-%m-%d'); break
            if fdate: break
        if not fdate: raise Bad('no date found above the header (expected e.g. 17-Nov-2026 in the top row)')
    rec = {}
    for r in rows[hi + 1:]:
        if not r or len(r) <= col['outlet_id'] or not r[col['outlet_id']]: continue
        sid = str(r[col['outlet_id']]).strip()
        if fdate: d = fdate
        else:
            v = r[col['Date']]
            if not v: continue
            d = v.strftime('%Y-%m-%d') if hasattr(v, 'strftime') else str(v)[:10]
        try: p = int(r[col['Planned']] or 0)
        except (TypeError, ValueError): p = 0
        sl = []
        for w in SLOTS:
            v = r[tcol[w]]
            try: sl.append(max(0, min(9, int(round(float(v or 0))))))
            except (TypeError, ValueError): sl.append(-1)
        name = re.sub(r'\s*\[[^\]]*\]\s*$', '', str(r[col['Store Name']] or sid)).replace('KPN Fresh ', 'KPN ').strip()
        rec[(d, sid)] = (p, sl, name, str(r[col['Cluster Name']] or '').strip())
    if not rec: raise Bad('no store rows below the header')
    return rec


# ---------- load ----------
files = sorted([f for f in glob.glob(os.path.join(src, '*.xls*')) + glob.glob(os.path.join(src, '*.csv'))
                if not os.path.basename(f).startswith('~$')], key=stamp)
if not files: sys.exit('No .xlsx/.csv file found in ' + src)

tm = {}
tm_path = os.path.join(HERE, 'config', 'tm_master.csv')
if os.path.exists(tm_path):
    for l in open(tm_path, encoding='utf-8-sig'):
        if l.strip() and ',' in l and not l.lower().startswith('cluster'):
            k, v = l.strip().split(',', 1); tm[k.strip()] = v.strip()

hist, names, clus = {}, {}, {}
newest = files[-1]
crit = []
before = None
for f in files:
    try: rec = parse(f)
    except Exception as e:
        msg = f'{os.path.basename(f)}: {e}'
        if f == newest: crit.append(msg)
        else: print('::warning::skipped ' + msg)
        continue
    if f == newest: before = dict(hist); new_rec = rec
    dates = {k[0] for k in rec}
    for k in [k for k in hist if k[0] in dates]: del hist[k]
    for (d, sid), (p, sl, n, c) in rec.items():
        hist[(d, sid)] = (p, sl)
        names.setdefault(sid, n)            # keep the first (oldest) name -> stable labels
        clus[sid] = c                       # latest cluster wins
    print(f'Read {os.path.basename(f)}: {len(rec)} rows, {min(dates)} .. {max(dates)}')

if crit:
    open(issues_path, 'w').write('CRITICAL — data.json NOT updated\n' + '\n'.join(crit) + '\n')
    for c in crit: print('::error::' + c)
    sys.exit(1)

# ---------- checks on the newest file ----------
warn = []
nd = sorted({k[0] for k in new_rec})
prev_dates = sorted({k[0] for k in before})
ist_yday = (dt.datetime.now(dt.timezone(dt.timedelta(hours=5, minutes=30))) - dt.timedelta(days=1)).strftime('%Y-%m-%d')
if len(nd) == 1 and nd[0] != ist_yday:
    warn.append(f'Newest file is dated {nd[0]}; yesterday (IST) is {ist_yday}')
if len(nd) == 1 and prev_dates and nd[0] in prev_dates:
    warn.append(f'{nd[0]} was already loaded — replaced with the newest file')
last7 = set(prev_dates[-7:])
active = {sid for (d, sid), (p, sl) in before.items() if d in last7 and any(sl)}
ids = {sid for (_, sid) in new_rec}
miss = sorted(active - ids)
if miss and len(nd) == 1: warn.append(f'{len(miss)} active store(s) missing: ' + ', '.join(names.get(s, s) for s in miss[:10]))
newids = sorted(ids - {sid for (_, sid) in before}) if before else []
if newids: warn.append('New outlet_id(s): ' + ', '.join(newids))
zero, noplan, over, badnum, unm = [], [], [], [], set()
P = O = 0
for (d, sid), (p, sl, n, c) in new_rec.items():
    if -1 in sl: badnum.append(n); continue
    if p <= 0: noplan.append(n); continue
    if not any(sl) and sid in active: zero.append(n)
    if max(sl) > p + 2: over.append(f'{n} ({max(sl)} vs plan {p})')
    if c not in tm: unm.add(c)
    if any(sl): P += 31 * p; O += sum(min(x, p) for x in sl)
if badnum: warn.append('Non-numeric till counts: ' + ', '.join(badnum[:8]))
if noplan: warn.append('Planned tills missing/0: ' + ', '.join(noplan[:8]))
if zero: warn.append(f'{len(zero)} active store(s) show 0 tills all day: ' + ', '.join(zero[:8]))
if over: warn.append('Till count far above plan: ' + ', '.join(over[:6]))
if unm: warn.append('COM not in config/tm_master.csv: ' + ', '.join(sorted(unm)))
day7 = []
for d in prev_dates[-7:]:
    pp = oo = 0
    for (dd, sid), (p, sl) in before.items():
        if dd == d and any(sl): pp += 31 * p; oo += sum(min(x, p) for x in sl)
    if pp: day7.append(oo / pp)
if P and day7 and len(nd) == 1:
    avg7 = statistics.mean(day7)
    if abs(O / P - avg7) > 0.10: warn.append(f'Overall manned {O / P:.1%} vs 7-day avg {avg7:.1%} — unusual swing, please verify')

# ---------- write data.json (same shape the dashboard expects) ----------
dates = sorted({k[0] for k in hist})
sids = sorted(names, key=lambda s: names[s])
di = {d: i for i, d in enumerate(dates)}; si = {s: i for i, s in enumerate(sids)}
rows = [[di[d], si[s], p, ''.join(map(str, [max(0, x) for x in sl]))] for (d, s), (p, sl) in sorted(hist.items())]
D = {'generated_at': dt.datetime.now(dt.timezone.utc).isoformat(timespec='seconds'),
     'source': os.path.basename(newest), 'dates': dates,
     'stores': [[s, names[s], clus[s]] for s in sids], 'tm': tm, 'rows': rows}
json.dump(D, open(out, 'w'), separators=(',', ':'))
print('Wrote', out, '| through', dates[-1], '| rows', len(rows), '| stores', len(sids),
      f'| newest-day overall {O / P:.1%}' if P else '')

if warn:
    open(issues_path, 'w').write(f'Data checks — {os.path.basename(newest)} ({", ".join(nd)})\n' + '\n'.join('- ' + w for w in warn) + '\n')
    for w in warn: print('::warning::' + w)
elif os.path.exists(issues_path):
    os.remove(issues_path)
