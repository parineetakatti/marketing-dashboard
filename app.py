import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Market Sales Briefing", layout="wide")

# ---------- fixed sheet layout ----------
HDR_ROW, SUB_ROW, DATA_START, COUNTRY_COL = 5, 6, 7, 2

COLS = {
    "growth": 55, "lifecycle": 79,
    "cyberthreat": 84,
    "dt_cult": 68, "dt_rep": 69, "dt_biz": 70,
    "theft": 20, "aging_u": 31, "aging_r": 32,
    "skill": 50, "cyber_cost": 48, "cyber_avail": 47,
    "costL2": 45, "costL3": 46, "availL1": 38, "availL2": 39, "availL3": 40,
    "recession": 57, "curfluct": 58,
    "elec_u": 3, "elec_r": 4, "backup_u": 7, "backup_r": 8,
    "4gu": 23, "5gu": 25, "proj": 27, "cloud": 37, "sme": 65,
    "embargo": 89,
    "psych": 92, "brand": 93, "colors": 95, "belief": 96,
}

def score(v):
    if v is None: return None
    v = str(v).lower()
    if "very high" in v or "significantly aging" in v or "very widely" in v: return 5
    if "moderate-high" in v: return 3.5
    if ("high" in v) or ("aging" in v and "not" not in v) or "widely" in v: return 4
    if "moderate" in v or "mixed" in v or "similar" in v: return 3
    if ("low" in v and "very" not in v) or "limited" in v or "modern" in v: return 2
    if "very low" in v or v.strip() in ("new","not widely","cheaper"): return 1
    return None

def sev_label(s):
    if s >= 5: return ("Very High", "vh")
    if s >= 4: return ("High", "h")
    return ("Moderate", "m")

@st.cache_data
def load_sheet(file):
    wb = load_workbook(file, data_only=True)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
    headers = {}
    for c in range(1, ws.max_column + 1):
        a = ws.cell(HDR_ROW, c).value
        b = ws.cell(SUB_ROW, c).value
        lbl = " ".join(str(x).strip() for x in [a, b] if x not in (None, ""))
        headers[c] = lbl if lbl else get_column_letter(c)
    rows = {}
    for r in range(DATA_START, ws.max_row + 1):
        nm = ws.cell(r, COUNTRY_COL).value
        if not nm: continue
        nm = str(nm).strip()
        if nm in rows: continue
        rec = {c: ws.cell(r, c).value for c in range(1, ws.max_column + 1)}
        rows[nm] = rec
    return headers, rows

def fears(d):
    out = []
    def addf(name, s, line):
        if s is None or s < 3: return
        lbl, cls = sev_label(s)
        out.append((name, lbl, cls, line))
    def addlow(name, s, line):
        if s is None or s > 2: return
        out.append((name, "Scarce", "h", line))
    addf("Cyber Security Threat", score(d.get(COLS["cyberthreat"])),
         "A breach here means severe reputation damage. Sell security as insurance — one incident dwarfs the subscription.")
    addf("Downtime → Reputation Loss", score(d.get(COLS["dt_rep"])),
         "An outage costs trust that's hard to rebuild here. Position uptime as brand protection.")
    addf("Downtime → Business Loss", score(d.get(COLS["dt_biz"])),
         "Downtime bleeds revenue directly. Quantify their hourly cost, then remove it.")
    addf("Downtime → Cultural Sensitivity", score(d.get(COLS["dt_cult"])),
         "Service interruptions carry outsized weight here. Reliability is a relationship issue, not just technical.")
    addf("Network Infrastructure Theft", score(d.get(COLS["theft"])),
         "Cable/fiber/battery theft drives repeat failures and truck-rolls. Our product cuts physical dependency.")
    addf("Aging Infrastructure", max(score(d.get(COLS["aging_u"])) or 0, score(d.get(COLS["aging_r"])) or 0) or None,
         "Constant patching of old kit. Modernization lets them leapfrog instead of nursing it.")
    addf("Skill Gap", score(d.get(COLS["skill"])),
         "Scarce talent to run complex infra. A managed/automated solution offsets the shortage.")
    addf("Cyber Expert Cost", score(d.get(COLS["cyber_cost"])),
         "Defenders cost a premium here. Automation lowers headcount dependency.")
    addf("Engineer Cost (L2/L3)", max(score(d.get(COLS["costL2"])) or 0, score(d.get(COLS["costL3"])) or 0) or None,
         "Every manual maintenance hour is expensive — reducing it is a direct opex win.")
    addf("Recession Risk", score(d.get(COLS["recession"])),
         "Budgets are defensive — frame as cost-saving, not new capex.")
    addf("Currency Fluctuation", score(d.get(COLS["curfluct"])),
         "Unpredictable costs. Emphasize stable, predictable total cost of ownership.")
    addf("Electricity Cuts", max(score(d.get(COLS["elec_u"])) or 0, score(d.get(COLS["elec_r"])) or 0) or None,
         "Power instability threatens uptime — position resilience as core, not optional.")
    addf("Power Backup Cost", max(score(d.get(COLS["backup_u"])) or 0, score(d.get(COLS["backup_r"])) or 0) or None,
         "Keeping kit powered is costly here — efficiency directly cuts their bill.")
    av = min(score(d.get(COLS["availL2"])) or 5, score(d.get(COLS["availL3"])) or 5)
    addlow("Engineers Hard to Find", av,
           "Thin local talent pool — our remote/managed model fills the gap they can't hire for.")
    addlow("Security Experts Scarce", score(d.get(COLS["cyber_avail"])),
           "Few local defenders — managed security covers what they can't staff.")
    return out

def greed(d):
    out = []
    def addg(name, line):
        out.append((name, line))
    g = d.get(COLS["growth"])
    try: gn = float(str(g).replace("%","")) if g not in (None,"") else None
    except: gn = None
    if gn is not None and gn >= 4:
        addg(f"Fast-Growing Economy ({g})", "Businesses are scaling fast and need infrastructure now — urgency favors you.")
    if (score(d.get(COLS["4gu"])) or 0) >= 4 or (score(d.get(COLS["5gu"])) or 0) >= 3:
        addg("Rapid 4G/5G Adoption", "Riding the connectivity wave — position the product as the efficient on-ramp.")
    if d.get(COLS["proj"]) and "rising" in str(d.get(COLS["proj"])).lower():
        addg("Heavy Planned Investment", "They're about to spend big here — be the build-out partner, not an afterthought.")
    if (score(d.get(COLS["cloud"])) or 0) >= 3:
        addg("Strong Cloud Adoption", "Cloud-ready buyers — slot in as the secure, efficient layer of their stack.")
    if (score(d.get(COLS["sme"])) or 5) <= 2:
        addg("Price-Sensitive Buyers (lead with ROI)", "Limited buying power — lead with cost-reduction and ROI, not premium features.")
    if not out:
        addg("Stable, Mature Market", "Lower risk, predictable buyer — compete on efficiency and reliability.")
    return out

def fmt_cost(v):
    if v in (None, ""): return None
    try: return f"{float(str(v).replace(',','')):,.0f}"
    except: return str(v)

# ---------- UI ----------
st.title("Market Sales Briefing")
st.caption("Upload the matrix → search a country → see the fears to press and the greed to ride.")

up = st.file_uploader("Upload the matrix (.xlsx)", type=["xlsx"])
if not up:
    st.info("Upload your Marketing Matrix Excel file to begin.")
    st.stop()

headers, rows = load_sheet(up)
names = sorted(rows.keys())

q = st.text_input("Search country", "")
filtered = [n for n in names if q.lower() in n.lower()] if q else names
picked = st.multiselect("Select one or more countries", filtered, default=filtered[:1] if filtered else [])

SEV_COLOR = {"vh": "#ef4444", "h": "#f97316", "m": "#f59e0b"}

for name in picked:
    d = rows[name]
    st.markdown(f"## {name}")
    g = d.get(COLS["growth"])
    if g: st.markdown(f"**GDP growth: {g}**  ·  _(World Bank)_")
    cost = fmt_cost(d.get(COLS["lifecycle"]))
    if cost:
        st.markdown(
            f"<div style='background:linear-gradient(90deg,#2a1a0d,#1a1410);border:1px solid #5a3a1a;"
            f"border-radius:10px;padding:12px 16px;margin:8px 0'>"
            f"<span style='color:#d89a5a;font-size:11px;text-transform:uppercase;letter-spacing:.5px;font-weight:700'>"
            f"Right now they bleed this just to keep the lights on</span><br>"
            f"<span style='color:#fbbf24;font-size:22px;font-weight:700'>~{cost}</span><br>"
            f"<span style='color:#9a8a72;font-size:12px'>per 1,000 sites / year, local currency — maintenance only. Our product shrinks this.</span></div>",
            unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### ⚠ Fears to press on")
        F = fears(d)
        if not F: st.write("_No major risk flags — sell on efficiency._")
        for nm, lbl, cls, line in F:
            col = SEV_COLOR.get(cls, "#f59e0b")
            st.markdown(
                f"<div style='border-left:3px solid {col};background:#1b222b;border-radius:8px;"
                f"padding:9px 12px;margin-bottom:9px'><b>{nm}</b> "
                f"<span style='background:{col};color:#111;font-size:10px;font-weight:700;"
                f"padding:1px 7px;border-radius:10px'>{lbl}</span><br>"
                f"<span style='color:#9aa7b5;font-size:12.5px'>{line}</span></div>",
                unsafe_allow_html=True)
    with c2:
        st.markdown("#### ▲ Greed to ride")
        for nm, line in greed(d):
            st.markdown(
                f"<div style='border-left:3px solid #22c55e;background:#1b222b;border-radius:8px;"
                f"padding:9px 12px;margin-bottom:9px'><b>{nm}</b><br>"
                f"<span style='color:#9aa7b5;font-size:12.5px'>{line}</span></div>",
                unsafe_allow_html=True)

    # ---------- Embargos ----------
    emb = d.get(COLS["embargo"])
    if emb not in (None, ""):
        is_clear = str(emb).strip().lower() in ("none", "no", "nil")
        ecol = "#3a4654" if is_clear else "#ef4444"
        st.markdown(
            f"<div style='border-left:4px solid {ecol};background:#1b222b;border-radius:8px;"
            f"padding:10px 14px;margin:6px 0'>"
            f"<span style='color:{ecol if not is_clear else '#9aa7b5'};font-size:11px;text-transform:uppercase;"
            f"letter-spacing:.5px;font-weight:700'>Embargos / Sanctions</span><br>"
            f"<span style='font-size:14px'>{emb}</span></div>",
            unsafe_allow_html=True)

    # ---------- Psychological & cultural ----------
    psych_fields = [("Human Psych & Negotiation", COLS["psych"]),
                    ("Brand Consciousness", COLS["brand"]),
                    ("Colors", COLS["colors"]),
                    ("Belief", COLS["belief"])]
    psych_rows = [(lbl, d.get(c)) for lbl, c in psych_fields if d.get(c) not in (None, "")]
    if psych_rows:
        st.markdown("#### 🧠 Psychological & Cultural")
        body = "".join(
            f"<div style='margin-bottom:7px'><b style='color:#a78bfa'>{lbl}:</b> "
            f"<span style='color:#cdd6e0;font-size:13px'>{val}</span></div>"
            for lbl, val in psych_rows)
        st.markdown(
            f"<div style='background:#1b222b;border-left:3px solid #8b5cf6;border-radius:8px;"
            f"padding:11px 14px;margin-bottom:8px'>{body}</div>",
            unsafe_allow_html=True)

    with st.expander("🔍 Deep look — every column for this country"):
        table = [{"Field": headers[c], "Value": str(d[c])} for c in sorted(d.keys())
                 if d[c] not in (None, "") and headers[c]]
        st.dataframe(pd.DataFrame(table), width='stretch', hide_index=True)
    st.divider()