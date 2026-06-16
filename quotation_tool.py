import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date
import json, os

# ── PDF ───────────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER

# ── Palette: Cloud & Cobalt ───────────────────────────────────────────────────
BG      = "#f8fafc"
SURFACE = "#f1f5f9"
CARD    = "#e2e8f0"
ACCENT  = "#1d4ed8"
TEXT    = "#0f172a"
MUTED   = "#64748b"
SUCCESS = "#15803d"
WHITE   = "#ffffff"
WARN    = "#b45309"

# ── Default product catalogue ─────────────────────────────────────────────────
TIERED_DEFAULTS = [
    (0,      1.40),
    (500,    1.25),
    (1000,   1.10),
    (5000,   0.95),
]

def _resource(filename):
    """Resolve path to a file sitting next to the exe (or script during dev)."""
    import sys
    if getattr(sys, 'frozen', False):
        # running as PyInstaller bundle — use exe's directory
        base = os.path.dirname(sys.executable)
    else:
        # running as plain .py script
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)

XLSX_FILE      = _resource("rawmaterial.xlsx")
CSV_FILE       = _resource("rwliste.csv")
CATALOGUE_FILE = "catalogue.json"
COUNTER_FILE   = "quote_counter.json"


def load_catalogue():
    """
    Load from rawmaterial.xlsx (preferred), then rwliste.csv, then catalogue.json.
    Returns a dict: { name -> [{"supplier","qty_kg","price_per_kg","option"}, ...] }
    """
    if os.path.exists(XLSX_FILE):
        result = import_catalogue_from_xlsx(XLSX_FILE)
        if result:
            return result
    if os.path.exists(CSV_FILE):
        result = import_catalogue_from_csv(CSV_FILE)
        if result:
            return result
    if os.path.exists(CATALOGUE_FILE):
        with open(CATALOGUE_FILE) as f:
            return json.load(f)
    return {}


def import_catalogue_from_xlsx(path):
    """
    Parse rawmaterial.xlsx — every row is a separate entry.
    Returns dict: { "idx:name" -> [{"supplier","qty_kg","price_per_kg","option"}] }
    Non-numeric prices are stored as-is in option so they're still visible.
    """
    try:
        import pandas as pd
        df = pd.read_excel(path, header=1)
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
                # store non-numeric price in option so it's still visible
                option = raw_price + (f"  |  {option}" if option else "")

            tier = {
                "supplier":      str(row["supplier"]).strip() if pd.notna(row["supplier"]) else "",
                "qty_kg":        str(row["qty_kg"]).strip()   if pd.notna(row["qty_kg"])   else "",
                "price_per_kg":  price,
                "price_valid":   price_valid,
                "option":        option,
            }
            # Use name as key — each row is its own entry stored as a single-tier list
            # Use a unique key to avoid merging: "i|name"
            key = f"{i}|{name}"
            catalogue[key] = [tier]
        return catalogue
    except Exception as e:
        print(f"XLSX load error: {e}")
        return {}


def import_catalogue_from_csv(path):
    """Legacy CSV loader — returns same dict structure."""
    import csv
    catalogue = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for i, row in enumerate(reader):
            # skip rows that are too short or have no name
            if len(row) < 2:
                continue
            try:
                # columns: id, name, supplier, qty_kg, price_per_kg, date, option
                name     = row[1].strip() if len(row) > 1 else ""
                supplier = row[2].strip() if len(row) > 2 else ""
                qty_kg   = row[3].strip() if len(row) > 3 else ""
                raw_price = row[4].strip() if len(row) > 4 else ""
                # skip date column (index 5)
                option   = row[6].strip() if len(row) > 6 else ""

                if not name:
                    continue

                try:
                    price = float(raw_price.replace(",", "."))
                    price_valid = True
                except ValueError:
                    price = 0.0
                    price_valid = False
                    if raw_price:
                        option = raw_price + (f"  |  {option}" if option else "")

                key = f"{i}|{name}"
                catalogue[key] = [{
                    "supplier": supplier,
                    "qty_kg": qty_kg,
                    "price_per_kg": price,
                    "price_valid": price_valid,
                    "option": option,
                }]
            except (ValueError, IndexError):
                continue
    return catalogue


def save_catalogue(cat):
    with open(CATALOGUE_FILE, "w") as f:
        json.dump(cat, f, indent=2)


def next_quote_number():
    n = 1
    if os.path.exists(COUNTER_FILE):
        with open(COUNTER_FILE) as f:
            n = json.load(f).get("last", 0) + 1
    with open(COUNTER_FILE, "w") as f:
        json.dump({"last": n}, f)
    return f"Q-{n:04d}"


# ─────────────────────────────────────────────────────────────────────────────
class QuotationApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("QuotePro – Food Supplements")
        self.geometry("1100x820")
        self.configure(bg=BG)
        self.resizable(True, True)

        self.catalogue   = load_catalogue()
        self.items       = []
        self.tax_rate    = tk.DoubleVar(value=0.0)
        self.discount    = tk.DoubleVar(value=0.0)
        self.margin_pct  = tk.DoubleVar(value=30.0)
        self.use_margin  = tk.BooleanVar(value=False)
        self.use_tiered  = tk.BooleanVar(value=False)
        self.tiers       = list(TIERED_DEFAULTS)   # [(min_bottles, mult), ...]
        self.quote_number = next_quote_number()

        self._build_ui()

    # ── helpers ───────────────────────────────────────────────────────────────
    def _entry(self, parent, width=14, **kw):
        e = tk.Entry(parent, bg=WHITE, fg=TEXT, insertbackground=TEXT,
                     selectbackground=ACCENT, selectforeground=WHITE,
                     relief="flat", font=("Helvetica", 10), width=width,
                     highlightthickness=1, highlightcolor=ACCENT,
                     highlightbackground=CARD, **kw)
        return e

    def _btn(self, parent, text, cmd, color=ACCENT, **kw):
        return tk.Button(parent, text=text, bg=color, fg=WHITE,
                         relief="flat", font=("Helvetica", 9, "bold"),
                         padx=8, pady=4, cursor="hand2", command=cmd, **kw)

    def _section(self, parent, title, pady=(10, 4)):
        tk.Label(parent, text=title, font=("Georgia", 11, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w", pady=pady)
        tk.Frame(parent, bg=ACCENT, height=2).pack(fill="x", pady=(0, 6))

    def _collapsible_section(self, parent, title, pady=(10, 4)):
        """Returns a body frame that toggles visibility when the header is clicked."""
        is_open = tk.BooleanVar(value=True)

        hdr = tk.Frame(parent, bg=BG, cursor="hand2")
        hdr.pack(fill="x", pady=pady)
        arrow = tk.Label(hdr, text="▼", font=("Helvetica", 9), bg=BG, fg=ACCENT)
        arrow.pack(side="left", padx=(0, 5))
        tk.Label(hdr, text=title, font=("Georgia", 11, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Frame(parent, bg=ACCENT, height=2).pack(fill="x", pady=(0, 6))

        body = tk.Frame(parent, bg=BG)
        body.pack(fill="x")

        def toggle(_=None):
            if is_open.get():
                body.pack_forget()
                arrow.config(text="▶")
            else:
                body.pack(fill="x")
                arrow.config(text="▼")
            is_open.set(not is_open.get())

        for w in [hdr, arrow] + hdr.winfo_children():
            w.bind("<Button-1>", toggle)
        # re-bind after children exist
        hdr.bind("<Button-1>", toggle)

        return body

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        hdr = tk.Frame(self, bg=CARD, pady=12, padx=20)
        hdr.pack(fill="x")
        tk.Label(hdr, text="QuotePro", font=("Georgia", 20, "bold"),
                 bg=CARD, fg=ACCENT).pack(side="left")
        tk.Label(hdr, text=f"Quote #{self.quote_number}   ·   {date.today():%d %b %Y}   ·   {len(self.catalogue)} ingredients loaded",
                 font=("Courier", 10), bg=CARD, fg=MUTED).pack(side="right")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=16, pady=12)

        # ── right sidebar (fixed, no scroll needed) ────────────────────────
        right = tk.Frame(body, bg=BG, width=280)
        right.pack(side="right", fill="y")

        # ── left: scrollable canvas ────────────────────────────────────────
        left_outer = tk.Frame(body, bg=BG)
        left_outer.pack(side="left", fill="both", expand=True, padx=(0, 12))

        canvas = tk.Canvas(left_outer, bg=BG, highlightthickness=0)
        vscroll = ttk.Scrollbar(left_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)

        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        left = tk.Frame(canvas, bg=BG)
        canvas_win = canvas.create_window((0, 0), window=left, anchor="nw")

        def _on_frame_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(e):
            canvas.itemconfig(canvas_win, width=e.width)

        left.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # mousewheel scroll — only scroll when pointer is over the canvas/left frame,
        # not when a popup (search dropdown) is open
        self._canvas = canvas

        def _on_mousewheel(e):
            if self._search_popup:
                return   # popup is open — let it handle its own scroll
            canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

        def _on_mousewheel_linux_up(e):
            if self._search_popup:
                return
            canvas.yview_scroll(-1, "units")

        def _on_mousewheel_linux_down(e):
            if self._search_popup:
                return
            canvas.yview_scroll(1, "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>",   _on_mousewheel_linux_up)
        canvas.bind_all("<Button-5>",   _on_mousewheel_linux_down)

        self._build_client(left)
        self._build_recipe(left)
        self._build_items(left)
        self._build_sidebar(right)

    # ── client ────────────────────────────────────────────────────────────────
    def _build_client(self, p):
        self._section(p, "Client Details")
        row = tk.Frame(p, bg=BG)
        row.pack(fill="x")
        fields = [("Client", "client"), ("Company", "company"),
                  ("Email",  "email"),  ("Phone",   "phone")]
        self._cfields = {}
        for i, (lbl, key) in enumerate(fields):
            col = tk.Frame(row, bg=BG)
            col.grid(row=0, column=i, padx=(0, 10), sticky="w")
            tk.Label(col, text=lbl, bg=BG, fg=MUTED,
                     font=("Helvetica", 9)).pack(anchor="w")
            e = self._entry(col, width=18)
            e.pack()
            self._cfields[key] = e

    # ── recipe builder ────────────────────────────────────────────────────────
    def _build_recipe(self, p):
        self._section(p, "Recipe Builder")
        self._recipe_ingredients = []

        outer = tk.Frame(p, bg=SURFACE, pady=10, padx=12)
        outer.pack(fill="x", pady=(0, 8))

        # ── top row: product name + bottle params ──────────────────────────
        top = tk.Frame(outer, bg=SURFACE)
        top.pack(fill="x", pady=(0, 8))

        tk.Label(top, text="Product name", bg=SURFACE, fg=MUTED,
                 font=("Helvetica", 9)).pack(side="left")
        self._recipe_name_e = self._entry(top, width=22)
        self._recipe_name_e.pack(side="left", padx=(4, 20))

        for lbl, attr, w, default in [
            ("Capsules / bottle", "_recipe_caps_e",    7, "60"),
            ("Bottles ordered",   "_recipe_bottles_e", 7, "100"),
        ]:
            tk.Label(top, text=lbl, bg=SURFACE, fg=MUTED,
                     font=("Helvetica", 9)).pack(side="left")
            e = self._entry(top, width=w)
            e.insert(0, default)
            e.pack(side="left", padx=(4, 16))
            e.bind("<KeyRelease>", lambda ev: self._recalc_recipe())
            setattr(self, attr, e)

        # ── ingredient search row ──────────────────────────────────────────
        ing_row = tk.Frame(outer, bg=CARD, pady=6, padx=8)
        ing_row.pack(fill="x", pady=(0, 4))

        # Search field
        tk.Label(ing_row, text="Search ingredient", bg=CARD, fg=MUTED,
                 font=("Helvetica", 9)).pack(side="left")
        self._search_var = tk.StringVar()
        self._search_e   = tk.Entry(ing_row, textvariable=self._search_var,
                                     bg=WHITE, fg=TEXT, insertbackground=TEXT,
                                     selectbackground=ACCENT, selectforeground=WHITE,
                                     relief="flat", font=("Helvetica", 10), width=28,
                                     highlightthickness=1, highlightcolor=ACCENT,
                                     highlightbackground=CARD)
        self._search_e.pack(side="left", padx=(4, 4))
        self._search_var.trace_add("write", self._on_search_change)
        self._search_e.bind("<Down>",   lambda e: self._focus_search_popup())
        self._search_e.bind("<Return>", lambda e: self._focus_search_popup())

        # Supplier badge
        self._supplier_var = tk.StringVar(value="")
        tk.Label(ing_row, textvariable=self._supplier_var, bg=CARD, fg=ACCENT,
                 font=("Helvetica", 9, "italic"), width=18,
                 anchor="w").pack(side="left", padx=(0, 12))

        tk.Label(ing_row, text="mg / capsule", bg=CARD, fg=MUTED,
                 font=("Helvetica", 9)).pack(side="left")
        self._rec_mg_e = self._entry(ing_row, width=8)
        self._rec_mg_e.pack(side="left", padx=(4, 12))

        tk.Label(ing_row, text="€ / kg", bg=CARD, fg=MUTED,
                 font=("Helvetica", 9)).pack(side="left")
        self._rec_price_e = self._entry(ing_row, width=8)
        self._rec_price_e.pack(side="left", padx=(4, 12))

        self._btn(ing_row, "＋ Add", self._add_recipe_ingredient,
                  color=ACCENT).pack(side="left")
        self._btn(ing_row, "📂 Import CSV", self._import_csv,
                  color=ACCENT).pack(side="left", padx=(8, 0))

        # ── dropdown listbox for search results ───────────────────────────
        self._search_popup  = None
        self._search_lb     = None
        self._selected_item = None

        # ── ingredients treeview ───────────────────────────────────────────
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("R.Treeview", background=WHITE, foreground=TEXT,
                         fieldbackground=WHITE, rowheight=24,
                         font=("Helvetica", 9))
        style.configure("R.Treeview.Heading", background=CARD, foreground=TEXT,
                         font=("Helvetica", 9, "bold"), relief="flat")
        style.map("R.Treeview", background=[("selected", ACCENT)],
                   foreground=[("selected", WHITE)])

        rec_cols = ("Ingredient", "Supplier", "mg / cap", "€ / kg", "Qty (kg)", "Option", "Cost / cap (€)")
        rec_frame = tk.Frame(outer, bg=WHITE, highlightthickness=1,
                              highlightbackground=CARD)
        rec_frame.pack(fill="x", pady=(4, 2))

        self._rec_tree = ttk.Treeview(rec_frame, columns=rec_cols, show="headings",
                                       style="R.Treeview", height=4)
        rec_widths  = [200, 100, 70, 75, 70, 140, 110]
        rec_anchors = ["w", "w", "e", "e", "e", "w", "e"]
        for col, w, a in zip(rec_cols, rec_widths, rec_anchors):
            self._rec_tree.heading(col, text=col)
            self._rec_tree.column(col, width=w, anchor=a)

        rec_vsb = ttk.Scrollbar(rec_frame, orient="vertical", command=self._rec_tree.yview)
        self._rec_tree.configure(yscrollcommand=rec_vsb.set)
        self._rec_tree.pack(side="left", fill="both", expand=True)
        rec_vsb.pack(side="right", fill="y")

        rm_row = tk.Frame(outer, bg=SURFACE)
        rm_row.pack(fill="x")
        self._btn(rm_row, "✕  Remove", self._remove_recipe_ingredient,
                  color=ACCENT).pack(side="left")

        # ── results row ───────────────────────────────────────────────────
        res = tk.Frame(outer, bg=SURFACE)
        res.pack(fill="x", pady=(8, 0))

        self._rec_cost_cap_var  = tk.StringVar(value="—")
        self._rec_cost_bot_var  = tk.StringVar(value="—")
        self._rec_total_var     = tk.StringVar(value="—")

        for lbl, var, col in [
            ("RM cost / capsule (€)", self._rec_cost_cap_var, TEXT),
            ("RM cost / bottle (€)",  self._rec_cost_bot_var, TEXT),
            ("Total RM cost (€)",     self._rec_total_var,    SUCCESS),
        ]:
            card = tk.Frame(res, bg=CARD, padx=10, pady=5)
            card.pack(side="left", padx=(0, 8))
            tk.Label(card, text=lbl, bg=CARD, fg=MUTED,
                     font=("Helvetica", 8)).pack()
            tk.Label(card, textvariable=var, bg=CARD, fg=col,
                     font=("Helvetica", 11, "bold")).pack()

        self._btn(res, "➕  Add to Quote", self._add_recipe_to_quote,
                  color=ACCENT).pack(side="right")

    # ── catalogue search / autocomplete ──────────────────────────────────────
    def _on_search_change(self, *_):
        query = self._search_var.get().strip().lower()
        self._close_search_popup()
        if not query or not self.catalogue:
            return
        matches = []
        for key, tiers in self.catalogue.items():
            # key format: "idx|name"
            name = key.split("|", 1)[1] if "|" in key else key
            if query in name.lower():
                tier = tiers[0]
                matches.append({"name": name, "key": key, **tier})
            if len(matches) >= 80:
                break
        if not matches:
            return
        self._open_search_popup(matches)

    def _open_search_popup(self, matches):
        x = self._search_e.winfo_rootx()
        y = self._search_e.winfo_rooty() + self._search_e.winfo_height() + 2

        popup = tk.Toplevel(self)
        popup.wm_overrideredirect(True)
        popup.configure(bg=WHITE)
        self._search_popup = popup

        border = tk.Frame(popup, bg=ACCENT, padx=1, pady=1)
        border.pack(fill="both", expand=True)
        inner = tk.Frame(border, bg=WHITE)
        inner.pack(fill="both", expand=True)

        hdr = tk.Frame(inner, bg=SURFACE, pady=4, padx=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text=f"{len(matches)} result{'s' if len(matches)!=1 else ''}",
                 bg=SURFACE, fg=ACCENT, font=("Helvetica", 9, "bold")).pack(side="left")
        tk.Label(hdr, text="↑↓ navigate   Enter / click to select   Esc to close",
                 bg=SURFACE, fg=MUTED, font=("Helvetica", 8)).pack(side="right")

        style = ttk.Style()
        style.configure("S.Treeview", background=WHITE, foreground=TEXT,
                         fieldbackground=WHITE, rowheight=26, font=("Helvetica", 9))
        style.configure("S.Treeview.Heading", background=CARD, foreground=TEXT,
                         font=("Helvetica", 9, "bold"), relief="flat")
        style.map("S.Treeview", background=[("selected", ACCENT)],
                   foreground=[("selected", WHITE)])

        cols = ("Ingredient", "Supplier", "Qty (kg)", "€ / kg", "Option")
        tv_frame = tk.Frame(inner, bg=WHITE)
        tv_frame.pack(fill="both", expand=True)

        tv = ttk.Treeview(tv_frame, columns=cols, show="headings",
                           style="S.Treeview", height=min(14, len(matches)))
        widths  = [260, 120, 75, 75, 130]
        anchors = ["w", "w", "e", "e", "w"]
        for col, w, a in zip(cols, widths, anchors):
            tv.heading(col, text=col)
            tv.column(col, width=w, anchor=a)

        vsb = ttk.Scrollbar(tv_frame, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        tv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        tv.tag_configure("odd",  background=WHITE)
        tv.tag_configure("even", background=SURFACE)
        tv.tag_configure("opt",  background="#fefce8")

        for i, m in enumerate(matches):
            has_opt = bool(m.get("option"))
            tag = "opt" if has_opt else ("even" if i % 2 == 0 else "odd")
            tv.insert("", "end", iid=str(i), tags=(tag,), values=(
                m["name"],
                m.get("supplier", ""),
                m.get("qty_kg", ""),
                f"€ {m['price_per_kg']:.2f}",
                m.get("option", ""),
            ))

        tv.selection_set("0")
        tv.focus("0")
        popup.geometry(f"670x{min(14, len(matches)) * 26 + 66}+{x}+{y}")
        self._search_lb = tv

        def on_pick(_=None):
            sel = tv.selection()
            if not sel:
                return
            try:
                m = matches[int(sel[0])]
            except (ValueError, IndexError):
                return
            self._close_search_popup()
            self._apply_tier(m["name"], m)

        def on_pick_click(e):
            iid = tv.identify_row(e.y)
            if iid:
                try:
                    m = matches[int(iid)]
                except (ValueError, IndexError):
                    return
                self._close_search_popup()
                self._apply_tier(m["name"], m)

        tv.bind("<Return>",          on_pick)
        tv.bind("<Double-Button-1>", on_pick_click)
        tv.bind("<Button-1>",        on_pick_click)
        tv.bind("<Escape>",          lambda e: self._close_search_popup())
        popup.bind("<FocusOut>",     lambda e: self.after(150, self._close_search_popup))

    def _apply_tier(self, name, tier):
        """Fill in ingredient fields from a selected tier."""
        self._selected_item = {"name": name, **tier}
        self._search_var.set(name)
        sup = tier.get("supplier", "")
        opt = tier.get("option", "")
        badge = sup + (f"  ⚠ {opt}" if opt else "")
        self._supplier_var.set(badge)
        self._rec_price_e.delete(0, "end")
        self._rec_price_e.insert(0, str(tier["price_per_kg"]))
        self._rec_mg_e.focus_set()

    def _focus_search_popup(self):
        """Move keyboard focus into the search results treeview."""
        if self._search_lb:
            self._search_lb.focus_set()
            children = self._search_lb.get_children()
            if children:
                self._search_lb.selection_set(children[0])
                self._search_lb.focus(children[0])

    def _close_search_popup(self):
        if self._search_popup:
            try:
                self._search_popup.destroy()
            except Exception:
                pass
            self._search_popup = None
            self._search_lb    = None

    def _import_csv(self):
        path = filedialog.askopenfilename(
            title="Import raw material CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            items = import_catalogue_from_csv(path)
            if not items:
                messagebox.showwarning("Empty", "No items found in CSV."); return
            self.catalogue = items
            # save as rwliste.csv so it auto-loads next time
            import shutil
            shutil.copy(path, CSV_FILE)
            messagebox.showinfo("Imported",
                                f"✅ {len(items)} items loaded.\n"
                                f"File saved as rwliste.csv — will auto-load next time.")
        except Exception as ex:
            messagebox.showerror("Import Error", str(ex))

    def _add_recipe_ingredient(self):
        name = self._search_var.get().strip()
        if not name:
            messagebox.showwarning("Missing", "Search and select an ingredient first.")
            return
        try:
            mg    = float(self._rec_mg_e.get())
            price = float(self._rec_price_e.get())
        except ValueError:
            messagebox.showwarning("Invalid", "mg and price must be numbers."); return

        supplier     = self._supplier_var.get() or (
            self._selected_item.get("supplier", "") if self._selected_item else "")
        qty_kg  = self._selected_item.get("qty_kg",  "") if self._selected_item else ""
        option  = self._selected_item.get("option",  "") if self._selected_item else ""
        cost_per_cap = (mg / 1_000_000) * price
        self._recipe_ingredients.append({
            "name": name, "supplier": supplier,
            "price_per_kg": price, "mg_per_capsule": mg,
            "qty_kg": qty_kg, "option": option,
        })
        self._rec_tree.insert("", "end", values=(
            name, supplier,
            f"{mg:g}",
            f"€ {price:.2f}",
            qty_kg,
            option,
            f"€ {cost_per_cap:.6f}",
        ))
        # reset search
        self._search_var.set("")
        self._supplier_var.set("")
        self._selected_item = None
        self._rec_mg_e.delete(0, "end")
        self._rec_price_e.delete(0, "end")
        self._recalc_recipe()

    def _remove_recipe_ingredient(self):
        for s in self._rec_tree.selection():
            idx = self._rec_tree.index(s)
            self._rec_tree.delete(s)
            self._recipe_ingredients.pop(idx)
        self._recalc_recipe()

    def _recalc_recipe(self):
        try:
            caps_bottle = float(self._recipe_caps_e.get())
            bottles     = float(self._recipe_bottles_e.get())
        except ValueError:
            self._rec_cost_cap_var.set("—")
            self._rec_cost_bot_var.set("—")
            self._rec_total_var.set("—")
            return

        cost_per_cap = sum(
            (ing["mg_per_capsule"] / 1_000_000) * ing["price_per_kg"]
            for ing in self._recipe_ingredients
        )
        cost_per_bot = cost_per_cap * caps_bottle
        total_rm     = cost_per_bot * bottles

        self._rec_cost_cap_var.set(f"{cost_per_cap:.6f}")
        self._rec_cost_bot_var.set(f"{cost_per_bot:.4f}")
        self._rec_total_var.set(f"{total_rm:.2f}")

    def _add_recipe_to_quote(self):
        if not self._recipe_ingredients:
            messagebox.showwarning("Empty Recipe",
                                   "Add at least one ingredient first."); return
        prod_name = self._recipe_name_e.get().strip() or "Custom Supplement"
        try:
            caps_bottle = int(float(self._recipe_caps_e.get()))
            bottles     = int(float(self._recipe_bottles_e.get()))
        except ValueError:
            messagebox.showwarning("Invalid", "Capsules/bottle and bottles must be numbers.")
            return

        cost_per_cap = sum(
            (ing["mg_per_capsule"] / 1_000_000) * ing["price_per_kg"]
            for ing in self._recipe_ingredients
        )
        cost_per_bot = cost_per_cap * caps_bottle
        desc = f"{prod_name} ({caps_bottle} caps/bottle, {len(self._recipe_ingredients)} ingredients)"

        self._add_item_row(
            desc=desc,
            qty=bottles,
            unit_price=cost_per_bot,
            unit_label="bottle",
            recipe=list(self._recipe_ingredients),   # store a copy
        )

    # ── line items table ──────────────────────────────────────────────────────
    def _build_items(self, p):
        self._section(p, "Line Items")

        inp = tk.Frame(p, bg=SURFACE, pady=6, padx=8)
        inp.pack(fill="x", pady=(0, 6))

        def ph_entry(w, ph):
            e = self._entry(inp, width=w)
            e.insert(0, ph)
            e.bind("<FocusIn>",  lambda ev, ee=e, t=ph: ee.delete(0, "end") if ee.get()==t else None)
            e.bind("<FocusOut>", lambda ev, ee=e, t=ph: ee.insert(0, t) if not ee.get() else None)
            e.pack(side="left", padx=3)
            return e

        self._idesc  = ph_entry(28, "Description")
        self._iqty   = ph_entry(6,  "Qty")
        self._iunit  = ph_entry(8,  "Unit")
        self._iprice = ph_entry(10, "Unit Price (€)")
        self._btn(inp, "＋ Add", self._add_manual_item).pack(side="left", padx=(6,0))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Q.Treeview", background=WHITE, foreground=TEXT,
                         fieldbackground=WHITE, rowheight=26,
                         font=("Helvetica", 9))
        style.configure("Q.Treeview.Heading", background=CARD, foreground=TEXT,
                         font=("Helvetica", 9, "bold"), relief="flat")
        style.map("Q.Treeview", background=[("selected", ACCENT)],
                   foreground=[("selected", WHITE)])

        cols = ("Description", "Qty", "Unit", "Unit Price", "Line Total", "Tiered Mult")
        tree_frame = tk.Frame(p, bg=WHITE, highlightthickness=1,
                               highlightbackground=CARD)
        tree_frame.pack(fill="x", pady=(0, 2))

        # show="tree headings" enables the native ▶/▼ expand arrow column
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="tree headings",
                                   style="Q.Treeview", height=8)
        # tree column (the arrow) — keep narrow
        self._tree.column("#0", width=18, minwidth=18, stretch=False)
        widths  = [262, 55, 60, 95, 95, 90]
        anchors = ["w", "e", "w", "e", "e", "e"]
        for col, w, a in zip(cols, widths, anchors):
            self._tree.heading(col, text=col)
            self._tree.column(col, width=w, anchor=a)

        # tag for ingredient child rows — indigo-tinted background
        self._tree.tag_configure("ingredient",
                                  background="#eef2ff", foreground="#3730a3",
                                  font=("Helvetica", 8))

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._btn(p, "✕  Remove Selected", self._remove_item,
                  color=ACCENT).pack(anchor="e", pady=(3, 0))

    def _tiered_multiplier(self, qty):
        """Return the multiplier that applies for a given bottle quantity."""
        if not self.use_tiered.get():
            return 1.0
        mult = 1.0
        for min_qty, m in sorted(self.tiers, key=lambda x: x[0]):
            if qty >= min_qty:
                mult = m
        return mult

    def _add_item_row(self, desc, qty, unit_price, unit_label="bottle", recipe=None):
        mult       = self._tiered_multiplier(qty)
        sell_price = unit_price * mult
        if self.use_margin.get():
            sell_price = unit_price / (1 - self.margin_pct.get() / 100)
            sell_price *= mult
        line_total = sell_price * qty

        self.items.append({
            "desc": desc, "qty": qty, "unit": unit_label,
            "unit_price": unit_price, "sell_price": sell_price,
            "mult": mult, "line_total": line_total,
            "recipe": recipe or [],
        })

        # insert parent row
        parent_iid = self._tree.insert("", "end", values=(
            desc, f"{qty:g}", unit_label,
            f"€ {sell_price:,.4f}", f"€ {line_total:,.2f}",
            f"×{mult:.2f}" if self.use_tiered.get() else "—",
        ))

        # insert ingredient child rows
        if recipe:
            for ing in recipe:
                cost_cap = (ing["mg_per_capsule"] / 1_000_000) * ing["price_per_kg"]
                qty_kg = ing.get("qty_kg", "")
                option = ing.get("option", "")
                detail = f"{ing['mg_per_capsule']:g} mg  ·  €{ing['price_per_kg']:.2f}/kg"
                if qty_kg:
                    detail += f"  ·  {qty_kg} kg"
                if option:
                    detail += f"  ⚠ {option}"
                self._tree.insert(parent_iid, "end", tags=("ingredient",), values=(
                    f"  ↳ {ing['name']}",
                    "",
                    ing.get("supplier", ""),
                    detail,
                    f"€ {cost_cap:.6f}",
                    "",
                ))

        self._refresh_totals()

    def _add_manual_item(self):
        desc  = self._idesc.get().strip()
        unit  = self._iunit.get().strip()
        if desc in ("Description", ""):
            messagebox.showwarning("Missing", "Enter a description."); return
        try:
            qty   = float(self._iqty.get())
            price = float(self._iprice.get().replace("€","").replace(",","").strip())
        except ValueError:
            messagebox.showwarning("Invalid", "Qty and Price must be numbers."); return
        self._add_item_row(desc, qty, price, unit_label=unit or "unit")
        for e, ph in [(self._idesc,"Description"),(self._iqty,"Qty"),
                      (self._iunit,"Unit"),(self._iprice,"Unit Price (€)")]:
            e.delete(0,"end"); e.insert(0, ph)

    def _remove_item(self):
        for s in self._tree.selection():
            # only remove top-level rows (parent == "")
            if self._tree.parent(s) == "":
                idx = [i for i, iid in enumerate(self._tree.get_children(""))
                       if iid == s]
                if idx:
                    self.items.pop(idx[0])
                self._tree.delete(s)
        self._refresh_totals()

    # ── sidebar ───────────────────────────────────────────────────────────────
    def _build_sidebar(self, p):
        self._section(p, "Pricing Rules")

        # Tax
        tk.Label(p, text="Tax Rate (%)", bg=BG, fg=MUTED,
                 font=("Helvetica", 9)).pack(anchor="w")
        tk.Spinbox(p, from_=0, to=50, increment=0.5, textvariable=self.tax_rate,
                   bg=SURFACE, fg=TEXT, insertbackground=TEXT, relief="flat",
                   font=("Helvetica", 10), width=8,
                   command=self._refresh_totals).pack(anchor="w", pady=(0,8))

        # Discount
        tk.Label(p, text="Discount on Total (%)", bg=BG, fg=MUTED,
                 font=("Helvetica", 9)).pack(anchor="w")
        disc_e = tk.Spinbox(p, from_=0, to=100, increment=1,
                             textvariable=self.discount,
                             bg=SURFACE, fg=TEXT, insertbackground=TEXT,
                             relief="flat", font=("Helvetica", 10), width=8,
                             command=self._refresh_totals)
        disc_e.pack(anchor="w", pady=(0,8))

        # Margin rule
        mf = tk.Frame(p, bg=BG)
        mf.pack(fill="x", pady=(0,4))
        tk.Checkbutton(mf, text="Apply Margin Rule", variable=self.use_margin,
                       bg=BG, fg=TEXT, selectcolor=CARD,
                       activebackground=BG, activeforeground=TEXT,
                       font=("Helvetica", 9),
                       command=self._rebuild_prices).pack(side="left")
        tk.Label(p, text="Target Margin (%)", bg=BG, fg=MUTED,
                 font=("Helvetica", 9)).pack(anchor="w")
        tk.Spinbox(p, from_=0, to=90, increment=1, textvariable=self.margin_pct,
                   bg=SURFACE, fg=TEXT, insertbackground=TEXT, relief="flat",
                   font=("Helvetica", 10), width=8,
                   command=self._rebuild_prices).pack(anchor="w", pady=(0,8))

        # Tiered pricing
        tf = tk.Frame(p, bg=BG)
        tf.pack(fill="x", pady=(0,4))
        tk.Checkbutton(tf, text="Tiered Pricing", variable=self.use_tiered,
                       bg=BG, fg=TEXT, selectcolor=CARD,
                       activebackground=BG, activeforeground=TEXT,
                       font=("Helvetica", 9),
                       command=self._rebuild_prices).pack(side="left")
        self._btn(tf, "Edit Tiers", self._edit_tiers,
                  color=ACCENT).pack(side="right")

        self._section(p, "Summary", pady=(14,4))

        self._sub_var   = tk.StringVar(value="€ 0.00")
        self._disc_var  = tk.StringVar(value="€ 0.00")
        self._tax_var   = tk.StringVar(value="€ 0.00")
        self._total_var = tk.StringVar(value="€ 0.00")

        for lbl, var, big in [("Subtotal",  self._sub_var,   False),
                               ("Discount",  self._disc_var,  False),
                               ("Tax",       self._tax_var,   False),
                               ("TOTAL",     self._total_var, True)]:
            row = tk.Frame(p, bg=CARD if big else BG)
            row.pack(fill="x", pady=(0, 1))
            tk.Label(row, text=lbl, bg=row["bg"], fg=MUTED if not big else SUCCESS,
                     font=("Helvetica", 9), padx=6).pack(side="left", pady=3)
            tk.Label(row, textvariable=var, bg=row["bg"],
                     fg=SUCCESS if big else TEXT,
                     font=("Helvetica", 13 if big else 10, "bold"),
                     padx=6).pack(side="right", pady=3)

        # Notes
        self._section(p, "Notes", pady=(12,4))
        self._notes = tk.Text(p, bg=SURFACE, fg=TEXT, insertbackground=TEXT,
                               relief="flat", font=("Helvetica", 9),
                               height=4, width=30, wrap="word",
                               highlightthickness=1, highlightcolor=ACCENT,
                               highlightbackground=CARD)
        self._notes.pack(fill="x")

        for txt, cmd, col in [
            ("📄  Export PDF",   self._export_pdf,  ACCENT),
            ("💾  Save JSON",    self._save_json,   CARD),
        ]:
            self._btn(p, txt, cmd, color=col).pack(fill="x", pady=(8,0))

    # ── totals ────────────────────────────────────────────────────────────────
    def _refresh_totals(self):
        subtotal = sum(i["line_total"] for i in self.items)
        disc_amt = subtotal * self.discount.get() / 100
        after    = subtotal - disc_amt
        tax_amt  = after * self.tax_rate.get() / 100
        total    = after + tax_amt
        self._sub_var.set(f"€ {subtotal:,.2f}")
        self._disc_var.set(f"- € {disc_amt:,.2f}")
        self._tax_var.set(f"€ {tax_amt:,.2f}")
        self._total_var.set(f"€ {total:,.2f}")

    def _rebuild_prices(self):
        """Recompute sell prices for all rows when rules change."""
        for item in self.items:
            mult = self._tiered_multiplier(item["qty"])
            sp   = item["unit_price"] * mult
            if self.use_margin.get():
                sp = item["unit_price"] / (1 - self.margin_pct.get() / 100) * mult
            item["sell_price"] = sp
            item["mult"]       = mult
            item["line_total"] = sp * item["qty"]

        # Refresh tree — re-insert all rows including children
        for row in self._tree.get_children():
            self._tree.delete(row)
        for item in self.items:
            parent_iid = self._tree.insert("", "end", values=(
                item["desc"], f"{item['qty']:g}", item["unit"],
                f"€ {item['sell_price']:,.4f}",
                f"€ {item['line_total']:,.2f}",
                f"×{item['mult']:.2f}" if self.use_tiered.get() else "—",
            ))
            for ing in item.get("recipe", []):
                cost_cap = (ing["mg_per_capsule"] / 1_000_000) * ing["price_per_kg"]
                qty_kg = ing.get("qty_kg", "")
                option = ing.get("option", "")
                detail = f"{ing['mg_per_capsule']:g} mg  ·  €{ing['price_per_kg']:.2f}/kg"
                if qty_kg:
                    detail += f"  ·  {qty_kg} kg"
                if option:
                    detail += f"  ⚠ {option}"
                self._tree.insert(parent_iid, "end", tags=("ingredient",), values=(
                    f"  ↳ {ing['name']}",
                    "",
                    ing.get("supplier", ""),
                    detail,
                    f"€ {cost_cap:.6f}",
                    "",
                ))
        self._refresh_totals()

    # ── tier editor ───────────────────────────────────────────────────────────
    def _edit_tiers(self):
        win = tk.Toplevel(self, bg=BG)
        win.title("Edit Tiered Pricing")
        win.geometry("360x300")

        tk.Label(win, text="Min Bottles   Multiplier",
                 bg=BG, fg=MUTED, font=("Courier", 10)).pack(pady=(12,4))

        frame = tk.Frame(win, bg=BG)
        frame.pack(fill="both", expand=True, padx=16)

        rows = []

        def render():
            for w in frame.winfo_children():
                w.destroy()
            rows.clear()
            for i, (mn, ml) in enumerate(self.tiers):
                r = tk.Frame(frame, bg=BG)
                r.pack(fill="x", pady=2)
                em = self._entry(r, width=10); em.insert(0, str(mn)); em.pack(side="left", padx=(0,8))
                el = self._entry(r, width=10); el.insert(0, str(ml)); el.pack(side="left")
                rows.append((em, el))
                self._btn(r, "✕", lambda i=i: (self.tiers.pop(i), render()),
                          color=SURFACE).pack(side="right")

        render()

        def add_tier():
            self.tiers.append((0, 1.0))
            render()

        def save_tiers():
            try:
                self.tiers = [(int(float(em.get())), float(el.get()))
                               for em, el in rows]
                self._rebuild_prices()
                win.destroy()
            except ValueError:
                messagebox.showwarning("Invalid", "Numbers only.", parent=win)

        bf = tk.Frame(win, bg=BG)
        bf.pack(pady=8)
        self._btn(bf, "＋ Add Tier", add_tier, color=CARD).pack(side="left", padx=4)
        self._btn(bf, "✔ Save",      save_tiers).pack(side="left", padx=4)

    # ── quote data ────────────────────────────────────────────────────────────
    def _quote_data(self):
        sub   = sum(i["line_total"] for i in self.items)
        disc  = sub * self.discount.get() / 100
        after = sub - disc
        tax   = after * self.tax_rate.get() / 100
        return {
            "quote_number": self.quote_number,
            "date": str(date.today()),
            "client": {k: e.get() for k, e in self._cfields.items()},
            "items": self.items,
            "subtotal": sub, "discount_pct": self.discount.get(),
            "discount_amt": disc, "tax_rate": self.tax_rate.get(),
            "tax_amt": tax, "total": after + tax,
            "notes": self._notes.get("1.0", "end").strip(),
        }

    # ── PDF export ────────────────────────────────────────────────────────────
    def _export_pdf(self):
        d = self._quote_data()
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile=f"{self.quote_number}.pdf")
        if not path:
            return

        doc = SimpleDocTemplate(path, pagesize=A4,
                                 leftMargin=18*mm, rightMargin=18*mm,
                                 topMargin=16*mm, bottomMargin=16*mm)
        styles = getSampleStyleSheet()
        H1 = ParagraphStyle("H1", fontSize=18, textColor=colors.HexColor("#1d4ed8"),
                              fontName="Helvetica-Bold", spaceAfter=2)
        H2 = ParagraphStyle("H2", fontSize=10, textColor=colors.HexColor("#64748b"),
                              fontName="Helvetica")
        NORM = styles["Normal"]
        NORM.textColor = colors.HexColor("#0f172a")
        RGHT = ParagraphStyle("R", parent=NORM, alignment=TA_RIGHT,
                               textColor=colors.HexColor("#0f172a"))
        BOLD = ParagraphStyle("B", parent=NORM, fontName="Helvetica-Bold",
                               textColor=colors.HexColor("#0f172a"))

        BG_C   = colors.HexColor("#1d4ed8")
        ACC_C  = colors.HexColor("#1d4ed8")
        LT_C   = colors.HexColor("#f1f5f9")
        TXT_C  = colors.HexColor("#0f172a")
        MUTED_C= colors.HexColor("#64748b")

        story = []

        # Header
        c = d["client"]
        hdr_data = [[
            Paragraph('<font color="white"><b>QuotePro</b></font>',
                      ParagraphStyle("H1w", fontSize=18, fontName="Helvetica-Bold")),
            Paragraph(
                f'<font color="#e2e8f0">Quote #{d["quote_number"]}<br/>'
                f'{date.today():%d %b %Y}</font>', RGHT)
        ]]
        hdr_tbl = Table(hdr_data, colWidths=["60%","40%"])
        hdr_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#1d4ed8")),
            ("TOPPADDING",  (0,0), (-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1), 8),
            ("LEFTPADDING", (0,0), (-1,-1), 10),
            ("RIGHTPADDING",(0,0), (-1,-1), 10),
        ]))
        story.append(hdr_tbl)
        story.append(Spacer(1, 8*mm))

        # Client info
        story.append(Paragraph("Bill To", ParagraphStyle("lbl", fontSize=8,
                                textColor=colors.HexColor("#64748b"), fontName="Helvetica")))
        story.append(Paragraph(
            f'<font color="#0f172a"><b>{c.get("client","")}</b> / {c.get("company","")}<br/>'
            f'{c.get("email","")}  {c.get("phone","")}</font>', NORM))
        story.append(Spacer(1, 6*mm))
        story.append(HRFlowable(width="100%", thickness=1,
                                 color=colors.HexColor("#1d4ed8")))
        story.append(Spacer(1, 4*mm))

        # Items table
        ING_BG  = colors.HexColor("#eef2ff")
        ING_TXT = colors.HexColor("#3730a3")

        col_heads = ["Description", "Qty", "Unit", "Unit Price", "Line Total"]
        tbl_data  = [[Paragraph(f'<font color="white"><b>{h}</b></font>', NORM)
                      for h in col_heads]]
        row_styles = []   # extra TableStyle commands per row

        data_row = 1  # track actual row index in tbl_data
        for it in d["items"]:
            clean_desc = it["desc"].replace("  ▶ click to expand", "")
            tbl_data.append([
                Paragraph(f'<b>{clean_desc}</b>', NORM),
                Paragraph(f'{it["qty"]:g}', RGHT),
                Paragraph(it["unit"], NORM),
                Paragraph(f'€ {it["sell_price"]:,.4f}', RGHT),
                Paragraph(f'€ {it["line_total"]:,.2f}', RGHT),
            ])
            data_row += 1

            # sub-rows for each ingredient
            for ing in it.get("recipe", []):
                cost_cap = (ing["mg_per_capsule"] / 1_000_000) * ing["price_per_kg"]
                qty_kg = ing.get("qty_kg", "")
                option = ing.get("option", "")
                detail = (f'({ing.get("supplier","")})  '
                          f'{ing["mg_per_capsule"]:g} mg  '
                          f'€{ing["price_per_kg"]:.2f}/kg')
                if qty_kg:
                    detail += f'  · {qty_kg} kg'
                if option:
                    detail += f'  ⚠ {option}'
                ing_label = (f'    ↳ {ing["name"]}  '
                             f'<font color="#64748b">{detail}</font>')
                tbl_data.append([
                    Paragraph(ing_label,
                              ParagraphStyle("ing", parent=NORM, fontSize=8,
                                             textColor=ING_TXT, leftIndent=10)),
                    Paragraph("", NORM),
                    Paragraph("", NORM),
                    Paragraph("", NORM),
                    Paragraph(f'€ {cost_cap:.6f}',
                              ParagraphStyle("ingr", parent=RGHT, fontSize=8,
                                             textColor=ING_TXT)),
                ])
                row_styles.append(
                    ("BACKGROUND", (0, data_row), (-1, data_row), ING_BG)
                )
                data_row += 1

        items_tbl = Table(tbl_data, colWidths=["40%","10%","10%","20%","20%"])
        base_style = [
            ("BACKGROUND", (0,0), (-1,0), BG_C),
            ("ROWBACKGROUNDS", (0,1), (-1,-1),
             [colors.white, colors.HexColor("#f1f5f9")]),
            ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
            ("TEXTCOLOR",   (0,1), (-1,-1), TXT_C),
            ("FONTSIZE",    (0,0), (-1,-1), 9),
            ("TOPPADDING",  (0,0), (-1,-1), 5),
            ("BOTTOMPADDING",(0,0),(-1,-1), 5),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING",(0,0),(-1,-1), 6),
            ("LINEBELOW",   (0,0), (-1,0), 1, ACC_C),
            ("LINEBELOW",   (0,1), (-1,-1), 0.5, colors.HexColor("#e2e8f0")),
        ] + row_styles
        items_tbl.setStyle(TableStyle(base_style))
        story.append(items_tbl)
        story.append(Spacer(1, 6*mm))

        # Totals
        def total_row(label, value, bold=False):
            fn = "Helvetica-Bold" if bold else "Helvetica"
            fc = colors.HexColor("#1d4ed8") if bold else colors.HexColor("#0f172a")
            return [
                Paragraph(f'<font name="{fn}" color="{fc.hexval()}">{label}</font>', RGHT),
                Paragraph(f'<font name="{fn}" color="{fc.hexval()}">{value}</font>', RGHT),
            ]

        totals = [
            total_row("Subtotal",  f'€ {d["subtotal"]:,.2f}'),
            total_row(f'Discount ({d["discount_pct"]:.1f}%)',
                      f'- € {d["discount_amt"]:,.2f}'),
            total_row(f'Tax ({d["tax_rate"]:.1f}%)',
                      f'€ {d["tax_amt"]:,.2f}'),
            total_row("TOTAL",     f'€ {d["total"]:,.2f}', bold=True),
        ]
        tot_tbl = Table(totals, colWidths=["75%","25%"])
        tot_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#f8fafc")),
            ("BACKGROUND", (0,3), (-1,3), colors.HexColor("#e2e8f0")),
            ("TOPPADDING",    (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING",  (0,0), (-1,-1), 8),
            ("LINEABOVE",     (0,3), (-1,3), 1, colors.HexColor("#1d4ed8")),
        ]))
        story.append(tot_tbl)

        if d["notes"]:
            story.append(Spacer(1, 8*mm))
            story.append(Paragraph("Notes", ParagraphStyle("lbl2", fontSize=9,
                          textColor=MUTED_C, fontName="Helvetica-Bold")))
            story.append(Paragraph(d["notes"], NORM))

        doc.build(story)
        messagebox.showinfo("PDF Exported", f"Saved to:\n{path}")

    # ── JSON save ─────────────────────────────────────────────────────────────
    def _save_json(self):
        d = self._quote_data()
        path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON", "*.json")],
            initialfile=f"{self.quote_number}.json")
        if path:
            with open(path, "w") as f:
                json.dump(d, f, indent=2, default=str)
            messagebox.showinfo("Saved", f"Saved to:\n{path}")


if __name__ == "__main__":
    app = QuotationApp()
    app.mainloop()
