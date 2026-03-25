import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="DR Text Generator", page_icon="📄", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans Thai', sans-serif;
}

.stApp {
    background-color: #0f1117;
    color: #e8eaf0;
}

h1, h2, h3 {
    font-family: 'IBM Plex Sans Thai', sans-serif;
    font-weight: 600;
}

.main-header {
    background: linear-gradient(135deg, #1a1f2e 0%, #0f1117 100%);
    border-bottom: 1px solid #2a2f3e;
    padding: 1.5rem 2rem;
    margin-bottom: 2rem;
}

.main-title {
    font-size: 1.6rem;
    font-weight: 600;
    color: #e8eaf0;
    letter-spacing: 0.02em;
}

.main-subtitle {
    font-size: 0.85rem;
    color: #6b7280;
    margin-top: 0.25rem;
    font-family: 'IBM Plex Mono', monospace;
}

.output-box {
    background-color: #1a1f2e;
    border: 1px solid #2a2f3e;
    border-radius: 8px;
    padding: 1.5rem;
    font-family: 'IBM Plex Sans Thai', sans-serif;
    font-size: 0.95rem;
    line-height: 2;
    color: #d1d5db;
    white-space: pre-wrap;
    min-height: 200px;
}

.exchange-header {
    color: #60a5fa;
    font-weight: 600;
    font-size: 1rem;
    margin-top: 0.5rem;
}

.section-header-label {
    color: #9ca3af;
    font-size: 0.75rem;
    font-family: 'IBM Plex Mono', monospace;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 0.5rem;
}

.stats-bar {
    display: flex;
    gap: 2rem;
    padding: 1rem 0;
    border-bottom: 1px solid #2a2f3e;
    margin-bottom: 1.5rem;
}

.stat-item {
    text-align: center;
}

.stat-value {
    font-size: 1.4rem;
    font-weight: 600;
    color: #60a5fa;
    font-family: 'IBM Plex Mono', monospace;
}

.stat-label {
    font-size: 0.75rem;
    color: #6b7280;
    margin-top: 0.1rem;
}

div[data-testid="stTabs"] button {
    font-family: 'IBM Plex Sans Thai', sans-serif;
    font-size: 0.85rem;
}

.stButton > button {
    background-color: #1e40af;
    color: white;
    border: none;
    border-radius: 6px;
    font-family: 'IBM Plex Sans Thai', sans-serif;
    font-weight: 500;
    padding: 0.5rem 1.2rem;
    transition: background-color 0.2s;
}

.stButton > button:hover {
    background-color: #2563eb;
    color: white;
}

.download-section {
    background-color: #111827;
    border: 1px solid #1f2937;
    border-radius: 6px;
    padding: 1rem;
    margin-top: 1rem;
}
</style>
""", unsafe_allow_html=True)

# ── Header ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <div class="main-title">📄 DR Text Generator</div>
    <div class="main-subtitle">Automated Thai-language DR filing text blocks from Excel template</div>
</div>
""", unsafe_allow_html=True)

# ── Exchange ordering config ──────────────────────────────────────────────────
EXCHANGE_ORDER = [
    "ฮ่องกง",
    "เซี่ยงไฮ้",
    "เซิ้นเจิ้น",
    "โตเกียว",
    "แนสแด็ก",
    "นิวยอร์ก",
    "นิวยอร์ก อาร์ก้า",
    "ปารีส",
]

def get_exchange_sort_key(exchange_name):
    for i, prefix in enumerate(EXCHANGE_ORDER):
        if exchange_name.strip().startswith(prefix):
            return i
    return 99

def parse_units(val):
    """Parse units value — handle bullet (•) as placeholder or numeric."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ["•", "-", ""]:
        return None
    try:
        v = float(s)
        if v == int(v):
            return int(v)
        return v
    except:
        return s

def parse_ratio(val):
    """Parse ratio — return formatted string with commas if needed."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ["•", "-", ""]:
        return None
    try:
        n = int(float(s))
        return f"{n:,}"
    except:
        return s

def load_data(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    rows = []

    # Single Stock sheet
    if "Single Stock" in xl.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name="Single Stock", header=1)
        df.columns = [str(c).strip() for c in df.columns]
        for _, row in df.iterrows():
            run = str(row.get("Run", "")).strip().upper()
            if run != "Y":
                continue
            full_name = str(row.get("Full company name", "")).strip()
            exchange = str(row.get("Exchange name", "")).strip()
            units = parse_units(row.get("Units"))
            ratio = parse_ratio(row.get("Ratio"))
            if not full_name or full_name == "nan":
                continue
            import re
            m = re.search(r'\("([^"]+)"\)', full_name)
            short_name = m.group(1) if m else str(row.get("Company name", "")).strip()
            rows.append({
                "type": "stock",
                "short_name": short_name,
                "full_name": full_name,
                "exchange": exchange,
                "units": units,
                "ratio": ratio,
            })

    # ETF sheet
    if "ETF" in xl.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name="ETF", header=1)
        df.columns = [str(c).strip() for c in df.columns]
        for _, row in df.iterrows():
            run = str(row.get("Run", "")).strip().upper()
            if run != "Y":
                continue
            etf_name = str(row.get("ETF Name", "")).strip()
            exchange = str(row.get("Exchange name", "")).strip()
            dr_ticker = str(row.get("DR Ticker", "")).strip()
            units = parse_units(row.get("Units"))
            ratio = parse_ratio(row.get("Ratio"))
            # Derive short name from DR Ticker (strip trailing "80" or "US80" etc.)
            short_name = dr_ticker.replace("80", "").replace("US", "").replace("LI", "").replace("LUS", "").strip() if dr_ticker else etf_name
            # For ETF short name, use the common ticker (GLD, etc.)
            # We'll store etf_name as full_name and derive short from DR ticker prefix
            rows.append({
                "type": "etf",
                "short_name": short_name,
                "full_name": etf_name,
                "exchange": exchange,
                "units": units,
                "ratio": ratio,
                "dr_ticker": dr_ticker,
            })

    rows.sort(key=lambda r: get_exchange_sort_key(r["exchange"]))
    return rows

def get_thai_exchange_header(exchange):
    """Return clean Thai exchange display name (before the parenthesis)."""
    s = str(exchange).split("(")[0].strip()
    # Normalise NYSE Arca variants
    if "อาร์ก้า" in s or "arca" in s.lower():
        return "นิวยอร์ก อาร์ก้า"
    return s

def format_units(n):
    """Format units number with Thai-style commas."""
    if n is None:
        return ""
    try:
        return f"{int(n):,}"
    except:
        return str(n)

# ── Output 1: Securities List ────────────────────────────────────────────────
def gen_output1(rows):
    lines = []
    ordered_keys, groups = group_by_exchange(rows)
    counter = 1
    for ex in ordered_keys:
        lines.append(f"ตลาดหลักทรัพย์{ex}")
        for r in groups[ex]:
            num = str(counter) + "."
            if r["type"] == "stock":
                line = f"{num.ljust(10)} หุ้นสามัญของ บริษัท {r['full_name']}"
            else:
                line = f"{num.ljust(10)} โครงการจัดการลงทุนต่างประเทศ {r['full_name']} (\"{r['short_name']}\")"
            lines.append(line)
            counter += 1
    return "\n".join(lines)

def group_by_exchange(rows):
    """Group rows by exchange in EXCHANGE_ORDER, collecting all rows per exchange."""
    groups = {}
    for r in rows:
        ex = get_thai_exchange_header(r["exchange"])
        groups.setdefault(ex, []).append(r)
    # Return in EXCHANGE_ORDER, then any remaining
    ordered_keys = []
    for prefix in EXCHANGE_ORDER:
        for ex in groups:
            if ex == prefix and ex not in ordered_keys:
                ordered_keys.append(ex)
    for ex in groups:
        if ex not in ordered_keys:
            ordered_keys.append(ex)
    return ordered_keys, groups

# ── Output 2: Units (restart per exchange) ──────────────────────────────────
def gen_output2(rows):
    lines = ["จำนวนหน่วยที่ขออนุญาตเสนอขาย:"]
    ordered_keys, groups = group_by_exchange(rows)
    for ex in ordered_keys:
        for i, r in enumerate(groups[ex], 1):
            num = str(i) + "."
            units_str = format_units(r["units"])
            lines.append(f"{num.ljust(10)} {r['short_name']} จำนวนไม่เกิน {units_str} ล้านหน่วย")
    return "\n".join(lines)

# ── Output 3: Ratios (sequential) ───────────────────────────────────────────
def gen_output3(rows):
    lines = []
    ordered_keys, groups = group_by_exchange(rows)
    counter = 1
    for ex in ordered_keys:
        for r in groups[ex]:
            num = str(counter) + "."
            ratio_str = r["ratio"] if r["ratio"] else ""
            lines.append(f"{num.ljust(10)} {r['short_name']} อัตราส่วน 1 ทรัพย์สินอ้างอิง : {ratio_str} DR")
            counter += 1
    return "\n".join(lines)

# ── Output 4: Units + Price (sequential) ────────────────────────────────────
def gen_output4(rows):
    lines = []
    ordered_keys, groups = group_by_exchange(rows)
    counter = 1
    for ex in ordered_keys:
        for r in groups[ex]:
            num = str(counter) + "."
            units_str = format_units(r["units"])
            lines.append(f"{num.ljust(10)} {r['short_name']} จำนวนไม่เกิน {units_str} ล้านหน่วย และราคาเป็นไปตามกลไกตลาดในเวลาที่เสนอขาย")
            counter += 1
    return "\n".join(lines)

# ── Output 5: Value (restart per exchange) ──────────────────────────────────
def gen_output5(rows):
    lines = ["มูลค่าที่คาดว่าจะเสนอขาย :"]
    ordered_keys, groups = group_by_exchange(rows)
    for ex in ordered_keys:
        for i, r in enumerate(groups[ex], 1):
            num = str(i) + "."
            lines.append(f"{num.ljust(10)} {r['short_name']} จำนวนไม่เกิน 10,000 ล้านบาท")
    return "\n".join(lines)

def make_txt_download(outputs):
    labels = [
        "=== OUTPUT 1: รายการหลักทรัพย์ ===",
        "=== OUTPUT 2: จำนวนหน่วยที่ขออนุญาตเสนอขาย ===",
        "=== OUTPUT 3: อัตราส่วน ===",
        "=== OUTPUT 4: จำนวนหน่วยและราคา ===",
        "=== OUTPUT 5: มูลค่าที่คาดว่าจะเสนอขาย ===",
    ]
    content = ""
    for label, text in zip(labels, outputs):
        content += label + "\n\n" + text + "\n\n\n"
    return content.encode("utf-8")

# ── Main UI ──────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload DR Filing Template (.xlsx)",
    type=["xlsx"],
    help="Upload the DR_Filing_Template Excel file with Single Stock and ETF sheets"
)

if uploaded:
    with st.spinner("Reading Excel file..."):
        rows = load_data(uploaded)

    if not rows:
        st.error("No valid rows found. Make sure Run column has 'Y' or 'N' values.")
    else:
        total = len(rows)
        stocks = sum(1 for r in rows if r["type"] == "stock")
        etfs = sum(1 for r in rows if r["type"] == "etf")
        exchanges = len(set(get_thai_exchange_header(r["exchange"]) for r in rows))

        st.markdown(f"""
        <div class="stats-bar">
            <div class="stat-item"><div class="stat-value">{total}</div><div class="stat-label">Total Securities</div></div>
            <div class="stat-item"><div class="stat-value">{stocks}</div><div class="stat-label">Single Stocks</div></div>
            <div class="stat-item"><div class="stat-value">{etfs}</div><div class="stat-label">ETFs</div></div>
            <div class="stat-item"><div class="stat-value">{exchanges}</div><div class="stat-label">Exchanges</div></div>
        </div>
        """, unsafe_allow_html=True)

        out1 = gen_output1(rows)
        out2 = gen_output2(rows)
        out3 = gen_output3(rows)
        out4 = gen_output4(rows)
        out5 = gen_output5(rows)
        outputs = [out1, out2, out3, out4, out5]

        tab_labels = [
            "1 · รายการหลักทรัพย์",
            "2 · จำนวนหน่วย",
            "3 · อัตราส่วน",
            "4 · หน่วย + ราคา",
            "5 · มูลค่า",
        ]

        tabs = st.tabs(tab_labels)
        for i, (tab, text) in enumerate(zip(tabs, outputs)):
            with tab:
                st.text_area(
                    label="",
                    value=text,
                    height=500,
                    key=f"output_{i}",
                    label_visibility="collapsed"
                )
                st.download_button(
                    label=f"⬇ Download Output {i+1} (.txt)",
                    data=text.encode("utf-8"),
                    file_name=f"dr_output_{i+1}.txt",
                    mime="text/plain",
                    key=f"dl_{i}"
                )

        st.divider()
        st.markdown("#### Download All Outputs")
        all_txt = make_txt_download(outputs)
        st.download_button(
            label="⬇ Download All 5 Outputs (.txt)",
            data=all_txt,
            file_name="dr_all_outputs.txt",
            mime="text/plain"
        )

else:
    st.info("👆 Upload the DR Filing Template Excel file to get started.")
    st.markdown("""
    **This tool generates 5 Thai-language text blocks:**
    - **Output 1** — Securities list grouped by exchange
    - **Output 2** — Units per security (numbering restarts per exchange)
    - **Output 3** — DR ratios (sequential numbering)
    - **Output 4** — Units + market price clause (sequential)
    - **Output 5** — Expected offering value at 10,000 ล้านบาท (restarts per exchange)
    """)
