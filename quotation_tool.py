import streamlit as st
import pandas as pd
import os, io, json
from datetime import date

# ── PDF ───────────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT

st.set_page_config(page_title="QuotePro", page_icon="💊", layout="wide")

# ── Palette: Cloud & Cobalt ───────────────────────────────────────────────────
ACCENT  = "#1d4ed8"
SUCCESS = "#15803d"
MUTED   = "#64748b"

TIERED_DEFAULTS = [(0, 1.40), (500, 1.25), (1000, 1.10), (5000, 0.95)]
XLSX_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rawmaterial.xlsx")


# ── Catalogue loading ────────────────────────────────────────────────────────
@st.cache_data
def load_catalogue():
    """Each row of rawmaterial.xlsx becomes its own catalogue entry."""
    if not os.path.exists(XLSX_FILE):
        return {}
    df = pd.read_excel(XLSX_FILE, header=1)
    df.columns = ["name", "qty_kg", "price_per_kg", "supplier", "option"]

    catalogue = {}
    for i, row in df.iterrows():
        if pd.isna(row["name"]):
            continue
        name = str(row["name"]).strip()
        if not name:
            continue

        raw_price = str(row["price_per_kg"]).strip() if pd.notna(row["price_per_kg"]) else ""
        try:
            price = float(raw_price.replace(",", "."))
            price_valid = True
        except (ValueError, AttributeError):
            price = 0.0
            price_valid = False

        option = str(row["option"]).strip() if pd.notna(row["option"]) else ""
        if not price_valid and raw_price:
            option = raw_price + (f"  |  {option}" if option else "")

        key = f"{i}|{name}"
        catalogue[key] = {
            "name": name,
            "supplier": str(row["supplier"]).strip() if pd.notna(row["supplier"]) else "",
            "qty_kg": str(row["qty_kg"]).strip() if pd.notna(row["qty_kg"]) else "",
            "price_per_kg": price,
            "price_valid": price_valid,
            "option": option,
        }
    return catalogue


def import_catalogue_from_csv(file_bytes):
    import csv, io as _io
    catalogue = {}
    text = file_bytes.decode("utf-8-sig")
    reader = csv.reader(_io.StringIO(text))
    for i, row in enumerate(reader):
        if len(row) < 2:
            continue
        try:
            name      = row[1].strip() if len(row) > 1 else ""
            supplier  = row[2].strip() if len(row) > 2 else ""
            qty_kg    = row[3].strip() if len(row) > 3 else ""
            raw_price = row[4].strip() if len(row) > 4 else ""
            option    = row[6].strip() if len(row) > 6 else ""
            if not name:
                continue
            try:
                price = float(raw_price.replace(",", "."))
                price_valid = True
            except ValueError:
                price, price_valid = 0.0, False
                if raw_price:
                    option = raw_price + (f"  |  {option}" if option else "")
            key = f"{i}|{name}"
            catalogue[key] = {
                "name": name, "supplier": supplier, "qty_kg": qty_kg,
                "price_per_kg": price, "price_valid": price_valid, "option": option,
            }
        except (ValueError, IndexError):
            continue
    return catalogue


# ── Session state init ───────────────────────────────────────────────────────
def init_state():
    defaults = {
        "catalogue": load_catalogue(),
        "items": [],                 # quote line items
        "recipe_ingredients": [],    # ingredients being built in current recipe
        "tax_rate": 0.0,
        "discount_pct": 0.0,
        "margin_pct": 30.0,
        "use_margin": False,
        "use_tiered": False,
        "tiers": list(TIERED_DEFAULTS),
        "quote_number": f"Q-{1:04d}",
        "notes": "",
        "client_name": "", "client_company": "", "client_email": "", "client_phone": "",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ── Pricing logic ────────────────────────────────────────────────────────────
def tiered_multiplier(qty):
    if not st.session_state.use_tiered:
        return 1.0
    mult = 1.0
    for min_qty, m in sorted(st.session_state.tiers, key=lambda x: x[0]):
        if qty >= min_qty:
            mult = m
    return mult


def calc_sell_price(unit_price, qty):
    mult = tiered_multiplier(qty)
    sell = unit_price * mult
    if st.session_state.use_margin:
        sell = (unit_price / (1 - st.session_state.margin_pct / 100)) * mult
    return sell, mult


def rebuild_prices():
    for item in st.session_state.items:
        sell, mult = calc_sell_price(item["unit_price"], item["qty"])
        item["sell_price"] = sell
        item["mult"] = mult
        item["line_total"] = sell * item["qty"]


def add_item(desc, qty, unit_price, unit="bottle", recipe=None):
    sell, mult = calc_sell_price(unit_price, qty)
    st.session_state.items.append({
        "desc": desc, "qty": qty, "unit": unit,
        "unit_price": unit_price, "sell_price": sell, "mult": mult,
        "line_total": sell * qty, "recipe": recipe or [],
    })


def cost_per_cap(ing):
    return (ing["mg_per_capsule"] / 1_000_000) * ing["price_per_kg"]


# ── PDF export ────────────────────────────────────────────────────────────────
def build_pdf(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                             leftMargin=18*mm, rightMargin=18*mm,
                             topMargin=16*mm, bottomMargin=16*mm)
    styles = getSampleStyleSheet()
    NORM = styles["Normal"]
    NORM.textColor = colors.HexColor("#0f172a")
    RGHT = ParagraphStyle("R", parent=NORM, alignment=TA_RIGHT, textColor=colors.HexColor("#0f172a"))
    BG_C = ACC_C = colors.HexColor("#1d4ed8")
    ING_BG, ING_TXT = colors.HexColor("#eef2ff"), colors.HexColor("#3730a3")
    MUTED_C = colors.HexColor("#64748b")

    story = []
    hdr_data = [[
        Paragraph('<font color="white"><b>QuotePro</b></font>',
                  ParagraphStyle("H1w", fontSize=18, fontName="Helvetica-Bold")),
        Paragraph(f'<font color="#e2e8f0">Quote #{data["quote_number"]}<br/>{date.today():%d %b %Y}</font>', RGHT)
    ]]
    hdr_tbl = Table(hdr_data, colWidths=["60%", "40%"])
    hdr_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BG_C),
        ("TOPPADDING", (0,0), (-1,-1), 8), ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("LEFTPADDING", (0,0), (-1,-1), 10), ("RIGHTPADDING", (0,0), (-1,-1), 10),
    ]))
    story += [hdr_tbl, Spacer(1, 8*mm)]

    c = data["client"]
    story.append(Paragraph("Bill To", ParagraphStyle("lbl", fontSize=8, textColor=MUTED_C, fontName="Helvetica")))
    story.append(Paragraph(
        f'<font color="#0f172a"><b>{c.get("name","")}</b> / {c.get("company","")}<br/>'
        f'{c.get("email","")}  {c.get("phone","")}</font>', NORM))
    story += [Spacer(1, 6*mm), HRFlowable(width="100%", thickness=1, color=ACC_C), Spacer(1, 4*mm)]

    col_heads = ["Description", "Qty", "Unit", "Unit Price", "Line Total"]
    tbl_data = [[Paragraph(f'<font color="white"><b>{h}</b></font>', NORM) for h in col_heads]]
    row_styles, data_row = [], 1

    for it in data["items"]:
        tbl_data.append([
            Paragraph(f'<b>{it["desc"]}</b>', NORM),
            Paragraph(f'{it["qty"]:g}', RGHT),
            Paragraph(it["unit"], NORM),
            Paragraph(f'€ {it["sell_price"]:,.4f}', RGHT),
            Paragraph(f'€ {it["line_total"]:,.2f}', RGHT),
        ])
        data_row += 1
        for ing in it.get("recipe", []):
            cc = cost_per_cap(ing)
            detail = f'({ing.get("supplier","")})  {ing["mg_per_capsule"]:g} mg  €{ing["price_per_kg"]:.2f}/kg'
            if ing.get("qty_kg"): detail += f'  · {ing["qty_kg"]} kg'
            if ing.get("option"): detail += f'  ⚠ {ing["option"]}'
            tbl_data.append([
                Paragraph(f'    ↳ {ing["name"]}  <font color="#64748b">{detail}</font>',
                          ParagraphStyle("ing", parent=NORM, fontSize=8, textColor=ING_TXT, leftIndent=10)),
                Paragraph("", NORM), Paragraph("", NORM), Paragraph("", NORM),
                Paragraph(f'€ {cc:.6f}', ParagraphStyle("ingr", parent=RGHT, fontSize=8, textColor=ING_TXT)),
            ])
            row_styles.append(("BACKGROUND", (0, data_row), (-1, data_row), ING_BG))
            data_row += 1

    items_tbl = Table(tbl_data, colWidths=["40%","10%","10%","20%","20%"])
    items_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), BG_C),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#f1f5f9")]),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("TEXTCOLOR", (0,1), (-1,-1), colors.HexColor("#0f172a")),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6), ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("LINEBELOW", (0,0), (-1,0), 1, ACC_C),
        ("LINEBELOW", (0,1), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
    ] + row_styles))
    story += [items_tbl, Spacer(1, 6*mm)]

    def trow(label, value, bold=False):
        fn = "Helvetica-Bold" if bold else "Helvetica"
        fc = colors.HexColor("#1d4ed8") if bold else colors.HexColor("#0f172a")
        return [Paragraph(f'<font name="{fn}" color="{fc.hexval()}">{label}</font>', RGHT),
                Paragraph(f'<font name="{fn}" color="{fc.hexval()}">{value}</font>', RGHT)]

    totals_tbl = Table([
        trow("Subtotal", f'€ {data["subtotal"]:,.2f}'),
        trow(f'Discount ({data["discount_pct"]:.1f}%)', f'- € {data["discount_amt"]:,.2f}'),
        trow(f'Tax ({data["tax_rate"]:.1f}%)', f'€ {data["tax_amt"]:,.2f}'),
        trow("TOTAL", f'€ {data["total"]:,.2f}', bold=True),
    ], colWidths=["75%","25%"])
    totals_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#f8fafc")),
        ("BACKGROUND", (0,3), (-1,3), colors.HexColor("#e2e8f0")),
        ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("LINEABOVE", (0,3), (-1,3), 1, ACC_C),
    ]))
    story.append(totals_tbl)

    if data["notes"]:
        story += [Spacer(1, 8*mm),
                  Paragraph("Notes", ParagraphStyle("lbl2", fontSize=9, textColor=MUTED_C, fontName="Helvetica-Bold")),
                  Paragraph(data["notes"], NORM)]

    doc.build(story)
    buffer.seek(0)
    return buffer


def get_quote_data():
    subtotal = sum(i["line_total"] for i in st.session_state.items)
    disc_amt = subtotal * st.session_state.discount_pct / 100
    after = subtotal - disc_amt
    tax_amt = after * st.session_state.tax_rate / 100
    return {
        "quote_number": st.session_state.quote_number,
        "client": {"name": st.session_state.client_name, "company": st.session_state.client_company,
                   "email": st.session_state.client_email, "phone": st.session_state.client_phone},
        "items": st.session_state.items,
        "subtotal": subtotal, "discount_pct": st.session_state.discount_pct, "discount_amt": disc_amt,
        "tax_rate": st.session_state.tax_rate, "tax_amt": tax_amt, "total": after + tax_amt,
        "notes": st.session_state.notes,
    }


# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
.stApp {{ background: #f8fafc; }}
div[data-testid="stHeader"] {{ background: transparent; }}
.quotepro-header {{
    background: {ACCENT}; padding: 1rem 1.5rem; border-radius: 12px;
    color: white; display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 1rem;
}}
.quotepro-title {{ font-size: 1.6rem; font-weight: 700; font-family: Georgia, serif; }}
.quotepro-meta {{ font-size: 0.85rem; opacity: 0.85; font-family: monospace; }}
.section-card {{
    background: white; border: 1px solid #e2e8f0; border-radius: 12px;
    padding: 1.25rem; margin-bottom: 1rem;
}}
.section-title {{ font-size: 1.1rem; font-weight: 700; color: #0f172a; margin-bottom: .75rem; font-family: Georgia, serif; }}
.ing-detail {{ background: #eef2ff; color: #3730a3; padding: 6px 10px; border-radius: 6px; font-size: 0.8rem; margin: 2px 0; }}
.summary-row {{ display: flex; justify-content: space-between; padding: 4px 0; font-size: 0.9rem; color: #475569; }}
.summary-total {{ display: flex; justify-content: space-between; padding: 8px 0; font-size: 1.15rem;
                   font-weight: 700; color: {ACCENT}; border-top: 1px solid #e2e8f0; margin-top: 4px; }}
button[kind="primary"] {{ background-color: {ACCENT} !important; border-color: {ACCENT} !important; }}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════ HEADER ══════════════════════════════════════
n_cat = len(st.session_state.catalogue)
st.markdown(f"""
<div class="quotepro-header">
  <div class="quotepro-title">QuotePro</div>
  <div class="quotepro-meta">{st.session_state.quote_number} · {date.today():%d %b %Y} · {n_cat} ingredients loaded</div>
</div>
""", unsafe_allow_html=True)

# client details
with st.container():
    c1, c2, c3, c4 = st.columns(4)
    st.session_state.client_name    = c1.text_input("Client Name", st.session_state.client_name)
    st.session_state.client_company = c2.text_input("Company",     st.session_state.client_company)
    st.session_state.client_email   = c3.text_input("Email",       st.session_state.client_email)
    st.session_state.client_phone   = c4.text_input("Phone",       st.session_state.client_phone)

main_col, side_col = st.columns([3, 1])

# ══════════════════════════════ MAIN COLUMN ═════════════════════════════════
with main_col:

    # ── Recipe Builder ──────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Recipe Builder</div>', unsafe_allow_html=True)

    rb1, rb2, rb3 = st.columns(3)
    product_name    = rb1.text_input("Product name", key="product_name", placeholder="e.g. Immune Booster")
    caps_per_bottle = rb2.number_input("Capsules / bottle", min_value=1, value=60, key="caps_per_bottle")
    bottles_ordered = rb3.number_input("Bottles ordered", min_value=1, value=100, key="bottles_ordered")

    st.markdown("###### Add ingredient")
    search_query = st.text_input("Search ingredient", key="ing_search", placeholder="Type to search…", label_visibility="collapsed")

    matches = []
    if search_query.strip():
        q = search_query.lower()
        matches = [v for v in st.session_state.catalogue.values() if q in v["name"].lower()][:50]

    selected_ing = None
    if matches:
        options = [f'{m["name"]}  —  {m["supplier"]}  ·  {m["qty_kg"]} kg  ·  '
                   + (f'€{m["price_per_kg"]:.2f}/kg' if m["price_valid"] else f'⚠ {m["option"]}')
                   for m in matches]
        idx = st.selectbox(f"{len(matches)} result(s)", range(len(options)),
                            format_func=lambda i: options[i], key="ing_select")
        selected_ing = matches[idx]
    elif search_query.strip():
        st.caption("No matches found.")

    ai1, ai2, ai3 = st.columns(3)
    mg_per_cap = ai1.number_input("mg / capsule", min_value=0.0, value=0.0, key="mg_per_cap")
    default_price = selected_ing["price_per_kg"] if (selected_ing and selected_ing["price_valid"]) else 0.0
    price_per_kg = ai2.number_input("€ / kg", min_value=0.0, value=float(default_price), key="price_per_kg")
    ai3.markdown("<br>", unsafe_allow_html=True)
    if ai3.button("＋ Add Ingredient", use_container_width=True):
        if not selected_ing:
            st.warning("Select an ingredient first.")
        elif mg_per_cap <= 0 or price_per_kg <= 0:
            st.warning("mg/capsule and €/kg must be greater than 0.")
        else:
            st.session_state.recipe_ingredients.append({
                "name": selected_ing["name"], "supplier": selected_ing["supplier"],
                "mg_per_capsule": mg_per_cap, "price_per_kg": price_per_kg,
                "qty_kg": selected_ing["qty_kg"], "option": selected_ing["option"],
            })
            st.rerun()

    # ingredients table
    if st.session_state.recipe_ingredients:
        df = pd.DataFrame(st.session_state.recipe_ingredients)
        df["cost_per_cap"] = df.apply(lambda r: (r["mg_per_capsule"]/1_000_000)*r["price_per_kg"], axis=1)
        display_df = df[["name", "supplier", "mg_per_capsule", "price_per_kg", "qty_kg", "option", "cost_per_cap"]].copy()
        display_df.columns = ["Ingredient", "Supplier", "mg/cap", "€/kg", "Qty(kg)", "Option", "Cost/cap (€)"]
        display_df["€/kg"] = display_df["€/kg"].map(lambda x: f"€ {x:.2f}")
        display_df["Cost/cap (€)"] = display_df["Cost/cap (€)"].map(lambda x: f"€ {x:.6f}")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        rm_cols = st.columns([1,1,1,5])
        if rm_cols[0].button("✕ Remove last"):
            st.session_state.recipe_ingredients.pop()
            st.rerun()
        if rm_cols[1].button("Clear all"):
            st.session_state.recipe_ingredients = []
            st.rerun()
    else:
        st.caption("No ingredients added yet.")

    total_cost_cap = sum(cost_per_cap(i) for i in st.session_state.recipe_ingredients)
    total_cost_bottle = total_cost_cap * caps_per_bottle
    total_rm = total_cost_bottle * bottles_ordered

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("RM cost / capsule", f"€ {total_cost_cap:.6f}")
    m2.metric("RM cost / bottle",  f"€ {total_cost_bottle:.4f}")
    m3.metric("Total RM cost",     f"€ {total_rm:.2f}")
    if m4.button("➕ Add to Quote", type="primary", use_container_width=True,
                 disabled=not st.session_state.recipe_ingredients):
        name = product_name.strip() or "Custom Supplement"
        desc = f"{name} ({caps_per_bottle} caps/bottle, {len(st.session_state.recipe_ingredients)} ingredients)"
        add_item(desc, bottles_ordered, total_cost_bottle, unit="bottle",
                  recipe=[dict(i) for i in st.session_state.recipe_ingredients])
        st.session_state.recipe_ingredients = []
        st.session_state.product_name = ""
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Line Items ───────────────────────────────────────────────────────────
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Line Items</div>', unsafe_allow_html=True)

    with st.expander("➕ Add manual item"):
        mc1, mc2, mc3, mc4 = st.columns([3,1,1,1])
        man_desc  = mc1.text_input("Description", key="man_desc")
        man_qty   = mc2.number_input("Qty", min_value=0.0, value=1.0, key="man_qty")
        man_unit  = mc3.text_input("Unit", value="pcs", key="man_unit")
        man_price = mc4.number_input("Unit Price (€)", min_value=0.0, value=0.0, key="man_price")
        if st.button("+ Add Item"):
            if man_desc and man_price > 0:
                add_item(man_desc, man_qty, man_price, unit=man_unit)
                st.rerun()

    if st.session_state.items:
        for idx, item in enumerate(st.session_state.items):
            row = st.columns([5, 1, 1, 1.3, 1.3, 0.6])
            row[0].markdown(f"**{item['desc']}**")
            row[1].write(f"{item['qty']:g}")
            row[2].write(item['unit'])
            row[3].write(f"€ {item['sell_price']:.4f}")
            row[4].markdown(f"**€ {item['line_total']:,.2f}**")
            if row[5].button("✕", key=f"rm_{idx}"):
                st.session_state.items.pop(idx)
                st.rerun()

            if item.get("recipe"):
                with st.expander(f"📋 {len(item['recipe'])} ingredients", expanded=False):
                    for ing in item["recipe"]:
                        cc = cost_per_cap(ing)
                        detail = f"{ing['mg_per_capsule']:g} mg · €{ing['price_per_kg']:.2f}/kg"
                        if ing.get("qty_kg"): detail += f" · {ing['qty_kg']} kg"
                        if ing.get("option"): detail += f" ⚠ {ing['option']}"
                        st.markdown(
                            f'<div class="ing-detail">↳ <b>{ing["name"]}</b> '
                            f'({ing.get("supplier","")}) — {detail} — '
                            f'<b>€ {cc:.6f}</b>/cap</div>', unsafe_allow_html=True)
            st.divider()
    else:
        st.caption("No items yet. Add one from the Recipe Builder or manually above.")

    st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════ SIDEBAR COLUMN ══════════════════════════════
with side_col:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Pricing Rules</div>', unsafe_allow_html=True)

    new_tax = st.number_input("Tax Rate (%)", min_value=0.0, max_value=50.0,
                               value=st.session_state.tax_rate, step=0.5, key="tax_rate_input")
    if new_tax != st.session_state.tax_rate:
        st.session_state.tax_rate = new_tax
        rebuild_prices()

    st.session_state.discount_pct = st.number_input(
        "Discount on Total (%)", min_value=0.0, max_value=100.0,
        value=st.session_state.discount_pct, key="discount_input")

    new_margin_on = st.checkbox("Apply Margin Rule", value=st.session_state.use_margin)
    if new_margin_on != st.session_state.use_margin:
        st.session_state.use_margin = new_margin_on
        rebuild_prices()

    if st.session_state.use_margin:
        new_margin = st.number_input("Target Margin (%)", min_value=0.0, max_value=90.0,
                                      value=st.session_state.margin_pct, key="margin_input")
        if new_margin != st.session_state.margin_pct:
            st.session_state.margin_pct = new_margin
            rebuild_prices()

    new_tiered_on = st.checkbox("Tiered Pricing", value=st.session_state.use_tiered)
    if new_tiered_on != st.session_state.use_tiered:
        st.session_state.use_tiered = new_tiered_on
        rebuild_prices()

    if st.session_state.use_tiered:
        with st.expander("Edit Tiers"):
            for i, (min_q, mult) in enumerate(st.session_state.tiers):
                tc1, tc2, tc3 = st.columns([2,2,1])
                new_min  = tc1.number_input("Min bottles", value=float(min_q), key=f"tier_min_{i}", label_visibility="collapsed")
                new_mult = tc2.number_input("Multiplier", value=float(mult), step=0.01, key=f"tier_mult_{i}", label_visibility="collapsed")
                if (new_min, new_mult) != (min_q, mult):
                    st.session_state.tiers[i] = (new_min, new_mult)
                    rebuild_prices()
                if tc3.button("×", key=f"tier_rm_{i}"):
                    st.session_state.tiers.pop(i)
                    rebuild_prices()
                    st.rerun()
            if st.button("+ Add tier"):
                st.session_state.tiers.append((0, 1.0))
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # summary
    qd = get_quote_data()
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Summary</div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="summary-row"><span>Subtotal</span><span>€ {qd['subtotal']:,.2f}</span></div>
    <div class="summary-row"><span>Discount ({qd['discount_pct']:.1f}%)</span><span>- € {qd['discount_amt']:,.2f}</span></div>
    <div class="summary-row"><span>Tax ({qd['tax_rate']:.1f}%)</span><span>€ {qd['tax_amt']:,.2f}</span></div>
    <div class="summary-total"><span>TOTAL</span><span>€ {qd['total']:,.2f}</span></div>
    """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # notes
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.session_state.notes = st.text_area("Notes", value=st.session_state.notes, height=100, label_visibility="visible")
    st.markdown('</div>', unsafe_allow_html=True)

    # export
    pdf_buffer = build_pdf(qd) if st.session_state.items else None
    st.download_button("📄 Export PDF", data=pdf_buffer or b"",
                        file_name=f"{st.session_state.quote_number}.pdf", mime="application/pdf",
                        use_container_width=True, type="primary",
                        disabled=not st.session_state.items)
    st.download_button("💾 Save JSON", data=json.dumps(qd, indent=2, default=str),
                        file_name=f"{st.session_state.quote_number}.json", mime="application/json",
                        use_container_width=True, disabled=not st.session_state.items)

    with st.expander("📂 Import different catalogue (CSV)"):
        uploaded = st.file_uploader("Upload CSV", type="csv", label_visibility="collapsed")
        if uploaded:
            new_cat = import_catalogue_from_csv(uploaded.read())
            if new_cat:
                st.session_state.catalogue = new_cat
                st.success(f"✅ {len(new_cat)} items loaded.")
                st.rerun()
            else:
                st.warning("No items found in CSV.")
