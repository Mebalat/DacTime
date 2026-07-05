"""
Courrier — application de rédaction de courriers
Fonctionnalités : onglets, thème sombre/clair, modèles, historique destinataires,
                  correcteur LanguageTool (clic droit), export PDF, impression.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font as tkfont
from datetime import datetime, date
import subprocess, os, re, threading, json

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from PIL import Image, ImageDraw, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False

try:
    from reportlab.pdfgen import canvas as rl_canvas
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

try:
    import win32print
    import win32ui
    import win32con
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ================================================================ CONFIG

EXPEDITEUR  = ["", "", ""]
TELEPHONE   = "0619680176"
TABS_DEST   = "\t" * 7
OUTPUT_DIR  = r"D:\Mes Documents"
HIST_FILE   = os.path.join(OUTPUT_DIR, ".courrier_historique.json")
CONFIG_DIR  = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "DacTime")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
LT_LOCAL    = "http://localhost:8081/v2/check"
LT_PUBLIC   = "https://api.languagetool.org/v2/check"

# ================================================================ THÈMES

THEMES = {
    "clair": {
        "bg": "#f7f6f2", "bg2": "#eceae4", "bg3": "#dedad2",
        "fg": "#1a1a1a", "fg2": "#555",    "fg3": "#999",
        "entry_bg": "white",   "entry_fg": "#1a1a1a",
        "text_bg":  "white",   "text_fg":  "#1a1a1a",
        "cursor": "black",     "sel_bg":   "#ffe0c0",
        "sep": "#bbb",
    },
    "sombre": {
        "bg": "#1e1e1e", "bg2": "#2a2a2a", "bg3": "#333",
        "fg": "#d4d4d4", "fg2": "#aaa",    "fg3": "#666",
        "entry_bg": "#2d2d2d", "entry_fg": "#d4d4d4",
        "text_bg":  "#252525", "text_fg":  "#d4d4d4",
        "cursor": "#d4d4d4",   "sel_bg":   "#4a3820",
        "sep": "#444",
    },
}

# ================================================================ MODÈLES

MODELES = {
    "Mise en demeure": {
        "objet": "Mise en demeure — [sujet]",
        "corps": (
            "Madame, Monsieur,\n\n"
            "Par la présente, je vous mets en demeure de [action attendue] "
            "dans un délai de 8 jours à compter de la réception de ce courrier.\n\n"
            "En effet, [exposé des faits].\n\n"
            "Sans réponse de votre part dans ce délai, je me verrai contraint "
            "d'engager les procédures judiciaires appropriées.\n\n"
            "Dans l'attente d'une résolution amiable,"
        ),
    },
    "Résiliation de contrat": {
        "objet": "Résiliation du contrat [référence]",
        "corps": (
            "Madame, Monsieur,\n\n"
            "Je vous informe par la présente de ma décision de résilier le contrat "
            "[référence] souscrit le [date], conformément aux conditions générales en vigueur.\n\n"
            "Je vous demande de bien vouloir prendre acte de cette résiliation "
            "et de me confirmer sa prise en compte par retour de courrier.\n\n"
            "En vous remerciant,"
        ),
    },
    "Relance impayé": {
        "objet": "Relance — Facture N°[xxx] impayée",
        "corps": (
            "Madame, Monsieur,\n\n"
            "Sauf erreur de ma part, la facture N°[référence] d'un montant de [montant] €, "
            "émise le [date], reste à ce jour impayée.\n\n"
            "Je vous serais reconnaissant de bien vouloir régulariser cette situation "
            "dans les meilleurs délais.\n\n"
            "En l'absence de règlement sous 8 jours, je me réserverai le droit "
            "d'engager une procédure de recouvrement.\n\n"
            "Cordialement,"
        ),
    },
    "Réclamation": {
        "objet": "Réclamation — [sujet]",
        "corps": (
            "Madame, Monsieur,\n\n"
            "Je me permets de vous contacter afin de vous faire part de mon "
            "mécontentement concernant [sujet].\n\n"
            "En effet, [exposé des faits et griefs].\n\n"
            "Je vous demande donc de bien vouloir [action attendue] "
            "dans les meilleurs délais.\n\n"
            "Dans l'attente de votre retour,"
        ),
    },
    "Demande de remboursement": {
        "objet": "Demande de remboursement — [référence]",
        "corps": (
            "Madame, Monsieur,\n\n"
            "Suite à [contexte], je vous adresse la présente afin d'obtenir "
            "le remboursement de la somme de [montant] €.\n\n"
            "[Justification de la demande].\n\n"
            "Je vous prie de bien vouloir procéder à ce remboursement par virement "
            "ou cheque a l'ordre de [votre nom].\n\n"
            "Dans l'attente de votre réponse,"
        ),
    },
    "Courrier administratif": {
        "objet": "[Objet du courrier]",
        "corps": (
            "Madame, Monsieur,\n\n"
            "J'ai l'honneur de vous adresser la présente afin de [objet de la démarche].\n\n"
            "[Corps du courrier]\n\n"
            "Je reste à votre disposition pour tout renseignement complémentaire.\n\n"
            "Veuillez agréer, Madame, Monsieur, "
            "l'expression de mes salutations distinguées,"
        ),
    },
}

# ================================================================ HISTORIQUE

def charger_historique():
    try:
        with open(HIST_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def sauvegarder_historique(hist):
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(HIST_FILE, "w", encoding="utf-8") as f:
            json.dump(hist[:10], f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def ajouter_historique(dest):
    hist = charger_historique()
    hist = [d for d in hist if d.get("nom") != dest.get("nom")]
    hist.insert(0, dest)
    sauvegarder_historique(hist)

def charger_config():
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def sauvegarder_config(data):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


VERSION              = "1.2"
LARGEUR_IMPRESSION   = 75   # chars Courier 10pt sur A4 avec marges 2.5cm


def _safe_unlink(path):
    try:
        os.unlink(path)
    except Exception:
        pass


def _lignes_impression(txt_widget):
    """Nombre de lignes telles qu'elles seront imprimées (wrap à 75 chars)."""
    import math
    contenu = txt_widget.get("1.0", "end-1c")
    total = 0
    for ligne in contenu.split("\n"):
        total += max(1, math.ceil(len(ligne) / LARGEUR_IMPRESSION)) if ligne else 1
    return total


# ================================================================ ONGLET

class CourrierTab(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=app.T["bg"])
        self.app          = app
        self._match_ranges = {}  # {txt_widget: [(start, end, match), ...]}
        self._build()

    def _build(self):
        app = self.app
        T   = app.T

        # Titre
        tk.Label(self, text="DacTime", font=("Courier New", 14, "bold"),
                 bg=T["bg"], fg=T["fg"], anchor="w").pack(fill="x", padx=20, pady=(12, 0))
        tk.Frame(self, height=1, bg=T["sep"]).pack(fill="x", padx=20, pady=(4, 8))

        # Adresses côte à côte
        addr_row = tk.Frame(self, bg=T["bg"])
        addr_row.pack(fill="x", padx=20, pady=(0, 6))

        # Expéditeur (gauche)
        exp_outer = tk.Frame(addr_row, bg=T["bg"])
        exp_outer.pack(side="left", fill="both", expand=True, padx=(0, 6))
        tk.Label(exp_outer, text="EXPÉDITEUR", font=("Courier New", 9, "bold"),
                 bg=T["bg"], fg=T["fg3"], anchor="w").pack(fill="x")
        exp_inner = tk.Frame(exp_outer, bg=T["entry_bg"], padx=10, pady=6)
        exp_inner.pack(fill="both", expand=True)
        exp_def = getattr(self.app, "_exp_defaut", EXPEDITEUR)
        self.exp_nom     = self._champ_col(exp_inner, "Nom",      0, exp_def[0])
        self.exp_adresse = self._champ_col(exp_inner, "Adresse",  1, exp_def[1])
        self.exp_cpville = self._champ_col(exp_inner, "CP+Ville", 2, exp_def[2])

        # Bouton historique expéditeurs
        btn_hist = tk.Button(exp_outer, text="choisir ▾",
                             font=("Courier New", 8), bg=T["bg"], fg=T["fg3"],
                             relief="flat", bd=0, cursor="hand2",
                             command=self._choisir_expediteur)
        btn_hist.pack(anchor="e", padx=2, pady=(2, 0))

        # Destinataire (droite)
        dest_outer = tk.Frame(addr_row, bg=T["bg"])
        dest_outer.pack(side="left", fill="both", expand=True, padx=(6, 0))
        tk.Label(dest_outer, text="DESTINATAIRE", font=("Courier New", 9, "bold"),
                 bg=T["bg"], fg=T["fg3"], anchor="w").pack(fill="x")
        dest_inner = tk.Frame(dest_outer, bg=T["entry_bg"], padx=10, pady=6)
        dest_inner.pack(fill="both", expand=True)
        self.nom     = self._champ_col(dest_inner, "Prénom Nom",           0)
        self.societe = self._champ_col(dest_inner, "Organisation (opt.)", 1)
        self.adresse = self._champ_col(dest_inner, "Adresse",             2)
        self.cpville = self._champ_col(dest_inner, "CP + Ville",          3)

        # Courrier
        c = self._bloc("COURRIER")
        self.d_date  = self._champ(c, "Date", 0, width=16)
        self.d_date.set(date.today().strftime("%d/%m/%Y"))
        self._btn_today = tk.Button(c, text="Aujourd'hui",
                  font=("Courier New", 9), bg=T["bg2"], fg=T["fg"],
                  relief="flat", bd=0, padx=8, pady=2, cursor="hand2",
                  command=self._today)
        self._btn_today.grid(row=0, column=2, padx=(6, 0))
        self._btn_cal = tk.Button(c, text="📅", font=("Segoe UI Emoji", 11), bg=T["bg2"],
                  relief="flat", bd=0, padx=4, pady=1, cursor="hand2",
                  command=self._ouvrir_calendrier)
        self._btn_cal.grid(row=0, column=3, padx=(2, 0))
        self.d_objet = self._champ(c, "Objet", 1)
        self.d_objet.trace_add("write", lambda *_: app._update_tab_title(self))

        # Corps — zone paginée (39 lignes page 1, 50 lignes pages suivantes)
        tk.Label(self, text="CORPS DU COURRIER", font=("Courier New", 9, "bold"),
                 bg=T["bg"], fg=T["fg3"], anchor="w").pack(fill="x", padx=20)

        canvas_frame = tk.Frame(self, bg=T["bg"])
        canvas_frame.pack(fill="both", expand=True, padx=20, pady=(0, 6))
        self._canvas_pages = tk.Canvas(canvas_frame, bg=T["bg"], highlightthickness=0)
        self._vsb_pages    = tk.Scrollbar(canvas_frame, orient="vertical",
                                          command=self._canvas_pages.yview)
        self._canvas_pages.configure(yscrollcommand=self._vsb_pages.set)
        self._vsb_pages.pack(side="right", fill="y")
        self._canvas_pages.pack(side="left", fill="both", expand=True)

        self._pages_inner = tk.Frame(self._canvas_pages, bg=T["bg"])
        self._pages_win   = self._canvas_pages.create_window(
            (0, 0), window=self._pages_inner, anchor="nw")
        self._pages_inner.bind("<Configure>",
            lambda e: self._canvas_pages.configure(
                scrollregion=self._canvas_pages.bbox("all")))
        self._canvas_pages.bind("<Configure>",
            lambda e: self._canvas_pages.itemconfig(self._pages_win, width=e.width))
        self._canvas_pages.bind("<MouseWheel>",
            lambda e: self._canvas_pages.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._corps_pages = []
        self._ajouter_bloc_page(1)

        _btn_frame = tk.Frame(self, bg=T["bg"])
        _btn_frame.pack(fill="x", padx=20, pady=(0, 4))
        self._btn_plus = tk.Button(_btn_frame, text="+ Nouvelle page",
            font=("Courier New", 9), bg=T["bg"], fg=T["fg3"],
            relief="flat", bd=0, cursor="hand2",
            command=self._ajouter_page)
        self._btn_plus.pack(anchor="w", padx=4, pady=(2, 2))

        # Référence interne aux blocs pour le thème
        self._blocs = []

    def _bloc(self, titre, expand=False):
        T     = self.app.T
        outer = tk.Frame(self, bg=T["bg"])
        outer.pack(fill="both" if expand else "x", expand=expand, padx=20, pady=(0, 6))
        lbl = tk.Label(outer, text=titre, font=("Courier New", 9, "bold"),
                       bg=T["bg"], fg=T["fg3"], anchor="w")
        lbl.pack(fill="x")
        inner = tk.Frame(outer, bg=T["entry_bg"], padx=10, pady=6)
        inner.pack(fill="both", expand=expand)
        return inner

    def _champ_col(self, parent, label, row, default=""):
        """Champ compact pour les blocs côte à côte (sans largeur fixe de label)."""
        T = self.app.T
        tk.Label(parent, text=label, font=("Courier New", 9), bg=T["entry_bg"],
                 fg=T["fg2"], anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        var = tk.StringVar(value=default)
        tk.Entry(parent, textvariable=var, font=("Courier New", 11),
                 relief="solid", bd=1, bg=T["entry_bg"], fg=T["entry_fg"],
                 insertbackground=T["cursor"]).grid(row=row, column=1, sticky="ew",
                                                    pady=2, padx=(6, 0))
        parent.columnconfigure(1, weight=1)
        return var

    def _champ(self, parent, label, row, width=40):
        T = self.app.T
        tk.Label(parent, text=label, font=("Courier New", 9), bg=T["entry_bg"],
                 fg=T["fg2"], anchor="w", width=26).grid(row=row, column=0, sticky="w", pady=2)
        var = tk.StringVar()
        tk.Entry(parent, textvariable=var, font=("Courier New", 11), width=width,
                 relief="solid", bd=1, bg=T["entry_bg"], fg=T["entry_fg"],
                 insertbackground=T["cursor"]).grid(row=row, column=1, sticky="ew",
                                                    pady=2, padx=(6, 0))
        parent.columnconfigure(1, weight=1)
        return var

    def _today(self):
        self.d_date.set(date.today().strftime("%d/%m/%Y"))

    def _ouvrir_calendrier(self):
        import calendar as _cal_mod
        # Fermer si déjà ouvert (toggle)
        if hasattr(self, "_cal_pop") and self._cal_pop and self._cal_pop.winfo_exists():
            self._cal_pop.destroy()
            return

        T = self.app.T
        try:
            d_sel = datetime.strptime(self.d_date.get(), "%d/%m/%Y").date()
        except ValueError:
            d_sel = date.today()

        selected  = [d_sel]
        nav_year  = tk.IntVar(value=d_sel.year)
        nav_month = tk.IntVar(value=d_sel.month)
        fmt_var   = tk.StringVar(value="court")

        MOIS_C = ["Janvier","Février","Mars","Avril","Mai","Juin",
                  "Juillet","Août","Septembre","Octobre","Novembre","Décembre"]
        MOIS_L = ["janvier","février","mars","avril","mai","juin",
                  "juillet","août","septembre","octobre","novembre","décembre"]
        JOURS_L = ["lundi","mardi","mercredi","jeudi","vendredi","samedi","dimanche"]
        JOURS_C = ["Lu","Ma","Me","Je","Ve","Sa","Di"]

        # ── Fenêtre sans décoration ──────────────────────────────────────
        pop = tk.Toplevel(self)
        self._cal_pop = pop
        pop.overrideredirect(True)
        pop.configure(bg=T["sep"])

        outer = tk.Frame(pop, bg=T["sep"], padx=1, pady=1)
        outer.pack()
        wrap = tk.Frame(outer, bg=T["bg"])
        wrap.pack()

        # ── En-tête bleu ────────────────────────────────────────────────
        hdr = tk.Frame(wrap, bg="#2a6099", padx=6, pady=5)
        hdr.pack(fill="x")

        btn_prev = tk.Button(hdr, text="‹", font=("Segoe UI", 13, "bold"),
                             bg="#2a6099", fg="white", activebackground="#1d4f7a",
                             activeforeground="white", relief="flat", bd=0,
                             cursor="hand2", padx=6)
        btn_prev.pack(side="left")

        lbl_nav = tk.Label(hdr, font=("Courier New", 10, "bold"),
                           bg="#2a6099", fg="white", cursor="hand2")
        lbl_nav.pack(side="left", expand=True)

        btn_next = tk.Button(hdr, text="›", font=("Segoe UI", 13, "bold"),
                             bg="#2a6099", fg="white", activebackground="#1d4f7a",
                             activeforeground="white", relief="flat", bd=0,
                             cursor="hand2", padx=6)
        btn_next.pack(side="right")

        # ── Grille ──────────────────────────────────────────────────────
        grid = tk.Frame(wrap, bg=T["bg"], padx=8, pady=4)
        grid.pack()

        for col, j in enumerate(JOURS_C):
            fg = "#e06060" if col >= 5 else T["fg3"]
            tk.Label(grid, text=j, font=("Courier New", 8, "bold"),
                     bg=T["bg"], fg=fg, width=3).grid(row=0, column=col, pady=(0, 2))

        day_btns = []
        for r in range(6):
            for c in range(7):
                b = tk.Button(grid, font=("Courier New", 9), width=3, height=1,
                              relief="flat", bd=0, cursor="hand2",
                              bg=T["bg"], fg=T["fg"],
                              activebackground="#2a6099", activeforeground="white")
                b.grid(row=r + 1, column=c, padx=1, pady=1)
                day_btns.append(b)

        # ── Format ──────────────────────────────────────────────────────
        fmt_row = tk.Frame(wrap, bg=T["bg2"], padx=8, pady=5)
        fmt_row.pack(fill="x")

        for val, txt in [("court", "05/07/2026"), ("moyen", "5 juillet"), ("long", "samedi 5 juillet")]:
            tk.Radiobutton(fmt_row, text=txt, variable=fmt_var, value=val,
                           font=("Courier New", 8), bg=T["bg2"], fg=T["fg2"],
                           selectcolor=T["bg3"], activebackground=T["bg2"],
                           relief="flat", bd=0).pack(side="left", padx=4)

        btn_auj = tk.Button(fmt_row, text="Aujourd'hui",
                            font=("Courier New", 8), bg=T["bg3"], fg=T["fg"],
                            relief="flat", bd=0, cursor="hand2", padx=6)
        btn_auj.pack(side="right")

        # ── Logique ─────────────────────────────────────────────────────
        def _valider(d):
            fmt = fmt_var.get()
            if fmt == "court":
                texte = d.strftime("%d/%m/%Y")
            elif fmt == "moyen":
                texte = f"{d.day} {MOIS_L[d.month-1]} {d.year}"
            else:
                texte = f"{JOURS_L[d.weekday()]} {d.day} {MOIS_L[d.month-1]} {d.year}"
            self.d_date.set(texte)
            pop.destroy()

        def _rafraichir():
            y, m = nav_year.get(), nav_month.get()
            lbl_nav.config(text=f"{MOIS_C[m-1]}  {y}")
            first_wd     = date(y, m, 1).weekday()
            days_in_month = _cal_mod.monthrange(y, m)[1]
            today = date.today()

            for i, btn in enumerate(day_btns):
                day_num = i - first_wd + 1
                btn.unbind("<Enter>"); btn.unbind("<Leave>")
                if 1 <= day_num <= days_in_month:
                    d = date(y, m, day_num)
                    is_sel  = (d == selected[0])
                    is_today = (d == today)
                    is_we   = (d.weekday() >= 5)
                    if is_sel:
                        bg, fg = "#2a6099", "white"
                    elif is_today:
                        bg, fg = T["sel_bg"], T["fg"]
                    elif is_we:
                        bg, fg = T["bg"], "#e06060"
                    else:
                        bg, fg = T["bg"], T["fg"]
                    btn.config(text=str(day_num), state="normal", bg=bg, fg=fg,
                               command=lambda d=d: _valider(d))
                    if not is_sel:
                        btn.bind("<Enter>", lambda e, b=btn: b.config(bg="#2a6099", fg="white"))
                        btn.bind("<Leave>", lambda e, b=btn, bg_=bg, fg_=fg: b.config(bg=bg_, fg=fg_))
                else:
                    btn.config(text="", state="disabled", bg=T["bg"],
                               disabledforeground=T["bg"], command=lambda: None)

        def _prev():
            m, y = nav_month.get(), nav_year.get()
            nav_month.set(12 if m == 1 else m - 1)
            if m == 1: nav_year.set(y - 1)
            _rafraichir()

        def _next():
            m, y = nav_month.get(), nav_year.get()
            nav_month.set(1 if m == 12 else m + 1)
            if m == 12: nav_year.set(y + 1)
            _rafraichir()

        def _aujourd_hui():
            nav_year.set(date.today().year)
            nav_month.set(date.today().month)
            selected[0] = date.today()
            _rafraichir()

        btn_prev.config(command=_prev)
        btn_next.config(command=_next)
        btn_auj.config(command=_aujourd_hui)

        # Clic sur le label mois/an → revenir à aujourd'hui
        lbl_nav.bind("<Button-1>", lambda e: _aujourd_hui())

        _rafraichir()

        # ── Fermer au clic extérieur ─────────────────────────────────────
        def _clic_ext(event):
            if not pop.winfo_exists(): return
            wx, wy = pop.winfo_rootx(), pop.winfo_rooty()
            ww, wh = pop.winfo_width(), pop.winfo_height()
            if not (wx <= event.x_root <= wx + ww and wy <= event.y_root <= wy + wh):
                pop.destroy()
        pop.bind_all("<Button-1>", _clic_ext, add="+")
        pop.bind("<Destroy>", lambda e: pop.unbind_all("<Button-1>"))

        # ── Positionnement intelligent ───────────────────────────────────
        self.app._popup_smart(pop)

    def get_titre(self):
        objet = self.d_objet.get().strip()
        return objet[:30] if objet else "Nouveau courrier"

    def valider(self):
        for label, var in [("Nom / Prénom", self.nom),
                            ("Adresse",     self.adresse),
                            ("CP + Ville",  self.cpville),
                            ("Objet",       self.d_objet)]:
            if not var.get().strip():
                messagebox.showerror("Champ manquant", f"Le champ « {label} » est obligatoire.")
                return False
        return True

    def build_txt(self):
        dest = [v for v in [
            self.nom.get().strip(),
            self.societe.get().strip(),
            self.adresse.get().strip(),
            self.cpville.get().strip(),
        ] if v]
        exp = [self.exp_nom.get().strip(),
               self.exp_adresse.get().strip(),
               self.exp_cpville.get().strip()]
        lignes = [l for l in exp if l]
        for l in dest:
            lignes.append(TABS_DEST + l)
        lignes.append("")
        if self.d_date.get().strip():
            lignes.append(self.d_date.get().strip())
        if self.d_objet.get().strip():
            lignes.append("Objet : " + self.d_objet.get().strip())
        lignes.append("")
        for i, txt in enumerate(self._corps_pages):
            corps = txt.get("1.0", "end-1c").strip()
            if i > 0:
                lignes.append("\f")
            lignes += corps.splitlines() if corps else [""]
        return "\n".join(lignes)

    def nom_fichier(self):
        slug = re.sub(r"[^a-zA-Z0-9]", "_", self.d_objet.get().strip() or "courrier")
        slug = re.sub(r"_+", "_", slug).strip("_")[:40]
        return f"{date.today().strftime('%Y-%m-%d')}_{slug}.dactime"

    def build_dactime(self):
        pages = [t.get("1.0", "end-1c") for t in self._corps_pages]
        return {
            "version":      1,
            "exp_nom":      self.exp_nom.get().strip(),
            "exp_adresse":  self.exp_adresse.get().strip(),
            "exp_cpville":  self.exp_cpville.get().strip(),
            "dest_nom":     self.nom.get().strip(),
            "dest_societe": self.societe.get().strip(),
            "dest_adresse": self.adresse.get().strip(),
            "dest_cpville": self.cpville.get().strip(),
            "date":         self.d_date.get().strip(),
            "objet":        self.d_objet.get().strip(),
            "corps":        pages[0] if pages else "",
            "pages":        pages,
        }

    def vider(self):
        for v in [self.nom, self.societe, self.adresse, self.cpville, self.d_objet]:
            v.set("")
        exp_def = getattr(self.app, "_exp_defaut", EXPEDITEUR)
        self.exp_nom.set(exp_def[0])
        self.exp_adresse.set(exp_def[1])
        self.exp_cpville.set(exp_def[2])
        self.d_date.set(date.today().strftime("%d/%m/%Y"))
        for txt in self._corps_pages:
            txt.delete("1.0", "end")
        self._effacer_erreurs()

    def _choisir_expediteur(self):
        hist = getattr(self.app, "_hist_exp", [])
        if not hist:
            messagebox.showinfo("Expediteur", "Aucun expediteur sauvegarde pour l'instant.\n"
                                "Enregistrez un courrier pour alimenter l'historique.")
            return
        T    = self.app.T
        menu = tk.Menu(self, tearoff=0, bg=T["bg2"], fg=T["fg"],
                       activebackground="#2a6099", activeforeground="white")
        for exp in hist:
            label = exp.get("exp_nom", "?")
            if exp.get("exp_cpville"):
                label += f"  —  {exp['exp_cpville']}"
            menu.add_command(label=label,
                             command=lambda e=exp: self._appliquer_expediteur(e))
        try:
            menu.tk_popup(self.app.winfo_pointerx(), self.app.winfo_pointery())
        finally:
            menu.grab_release()

    def _appliquer_expediteur(self, exp):
        self.exp_nom.set(exp.get("exp_nom", ""))
        self.exp_adresse.set(exp.get("exp_adresse", ""))
        self.exp_cpville.set(exp.get("exp_cpville", ""))
        if exp.get("telephone"):
            self.app._tel_defaut = exp["telephone"]

    def remplir_depuis_historique(self, dest):
        self.nom.set(dest.get("nom", ""))
        self.societe.set(dest.get("societe", ""))
        self.adresse.set(dest.get("adresse", ""))
        self.cpville.set(dest.get("cpville", ""))

    # -------------------------------------------------------- pagination

    def _ajouter_bloc_page(self, num):
        """Crée un widget Text pour la page num dans le canvas scrollable."""
        T = self.app.T
        if num > 1:
            tk.Frame(self._pages_inner, height=1, bg=T["sep"]).pack(
                fill="x", pady=(12, 0))
            tk.Label(self._pages_inner, text=f"PAGE {num}",
                     font=("Courier New", 8, "bold"),
                     bg=T["bg"], fg=T["fg3"]).pack(anchor="w", pady=(4, 2))
        frame = tk.Frame(self._pages_inner, bg=T["entry_bg"])
        frame.pack(fill="x")
        # NE PAS MODIFIER — calibré et validé à l'impression physique.
        # Page 1 = 48 lignes (en-tête + corps tient sur 1 feuille notepad).
        # Page 2+ = 54 lignes (feuille pleine, sans débordement page 3).
        h = 48 if num == 1 else 54
        max_lignes = h
        txt = tk.Text(frame, height=h,
                      font=(self.app._font_family.get(), self.app._font_size.get()),
                      wrap="word", relief="flat",
                      bg=T["text_bg"], fg=T["text_fg"],
                      insertbackground=T["cursor"],
                      selectbackground=T["sel_bg"],
                      padx=6, pady=6, undo=True)
        txt.tag_configure("error",  underline=True, foreground="#cc2200")
        txt.tag_configure("trouve", background="#ffe080")
        txt.pack(fill="both", expand=True)

        txt._page_num   = num
        txt._max_lignes = max_lignes

        def _on_return(event, _txt=txt, _max=max_lignes, _n=num):
            if _lignes_impression(_txt) >= _max:
                self.app._statut.set(
                    f"PAGE {_n} PLEINE ({_max} lignes) — cliquez '+ Nouvelle page'")
                return "break"

        def _on_paste(event, _txt=txt, _max=max_lignes, _n=num):
            def _tronquer():
                import math
                lignes = _txt.get("1.0", "end-1c").split("\n")
                total, nb_garder = 0, len(lignes)
                for i, l in enumerate(lignes):
                    cout = max(1, math.ceil(len(l) / LARGEUR_IMPRESSION)) if l else 1
                    if total + cout > _max:
                        nb_garder = i
                        break
                    total += cout
                if nb_garder < len(lignes):
                    _txt.delete("1.0", "end")
                    _txt.insert("1.0", "\n".join(lignes[:nb_garder]))
                    self.app._statut.set(f"Page {_n} : texte tronqué à {_max} lignes")
            _txt.after(1, _tronquer)

        txt.bind("<Return>",  _on_return)
        txt.bind("<<Paste>>", _on_paste)
        txt.bind("<Button-3>",     self._on_right_click)
        txt.bind("<KeyRelease>",   lambda e: self.app._update_statusbar())
        txt.bind("<ButtonRelease>", lambda e: self.app._update_statusbar())
        txt.bind("<MouseWheel>",
            lambda e: self._canvas_pages.yview_scroll(int(-1*(e.delta/120)), "units"))
        if num == 1:
            self.corps = txt
        self._corps_pages.append(txt)
        return txt

    def _ajouter_page(self):
        num = len(self._corps_pages) + 1
        self._ajouter_bloc_page(num)
        txt = self._corps_pages[-1]
        txt.focus_set()
        def _scroll_to_top():
            self._canvas_pages.update_idletasks()
            total_h = self._pages_inner.winfo_height()
            if total_h > 0:
                y_top = txt.winfo_rooty() - self._pages_inner.winfo_rooty()
                self._canvas_pages.yview_moveto(max(0, y_top / total_h))
        self._canvas_pages.after(150, _scroll_to_top)

    # -------------------------------------------------------- LT (clic droit)

    def _on_right_click(self, event):
        txt = event.widget  # widget Text qui a reçu le clic (page 1, 2, ...)
        pos = txt.index(f"@{event.x},{event.y}")
        hit = None
        for start, end, match in self._match_ranges.get(txt, []):
            if (txt.compare(pos, ">=", start) and
                    txt.compare(pos, "<", end)):
                hit = (start, end, match)
                break

        T    = self.app.T
        menu = tk.Menu(self, tearoff=0, bg=T["bg2"], fg=T["fg"],
                       activebackground="#2a6099", activeforeground="white")

        if hit:
            start, end, match = hit
            repls = [r["value"] for r in match.get("replacements", [])][:6]
            if repls:
                for s in repls:
                    menu.add_command(
                        label=s, font=("Courier New", 11, "bold"),
                        foreground="#4a9fd4",
                        command=lambda v=s, a=start, b=end: self._corriger(txt, a, b, v))
            else:
                menu.add_command(label=match.get("message", "Erreur"), state="disabled")
            menu.add_separator()
            menu.add_command(label="Ignorer",
                             command=lambda a=start, b=end: self._ignorer(txt, a, b))
            menu.add_separator()

        menu.add_command(label="Couper",
                         command=lambda: txt.event_generate("<<Cut>>"))
        menu.add_command(label="Copier",
                         command=lambda: txt.event_generate("<<Copy>>"))
        menu.add_command(label="Coller",
                         command=lambda: txt.event_generate("<<Paste>>"))
        menu.add_separator()
        menu.add_command(label="Sélectionner tout",
                         command=lambda: txt.tag_add(tk.SEL, "1.0", "end"))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def afficher_erreurs(self, txt, texte, matches):
        self._effacer_erreurs_widget(txt)
        self._match_ranges.setdefault(txt, [])
        for m in matches:
            s = self._offset_idx(texte, m["offset"])
            e = self._offset_idx(texte, m["offset"] + m["length"])
            txt.tag_add("error", s, e)
            self._match_ranges[txt].append((s, e, m))

    def _offset_idx(self, texte, offset):
        avant  = texte[:offset]
        lignes = avant.split("\n")
        return f"{len(lignes)}.{len(lignes[-1])}"

    def _effacer_erreurs(self):
        for txt in list(self._match_ranges):
            txt.tag_remove("error", "1.0", "end")
        self._match_ranges = {}

    def _effacer_erreurs_widget(self, txt):
        txt.tag_remove("error", "1.0", "end")
        self._match_ranges.pop(txt, None)

    def _corriger(self, txt, start, end, valeur):
        txt.delete(start, end)
        txt.insert(start, valeur)
        self.app._verifier()

    def _ignorer(self, txt, start, end):
        txt.tag_remove("error", start, end)
        self._match_ranges[txt] = [(s, e, m) for s, e, m in self._match_ranges.get(txt, [])
                                   if not (s == start and e == end)]
        nb = sum(len(v) for v in self._match_ranges.values())
        self.app._statut.set(
            f"{nb} erreur{'s' if nb > 1 else ''} restante{'s' if nb > 1 else ''}"
            if nb else "")


# ================================================================ NOTEBOOK FERMABLE

class ClosableNotebook(ttk.Notebook):
    """ttk.Notebook avec bouton × sur chaque onglet (PIL requis)."""
    _style_done = False
    _img_n      = None
    _img_a      = None

    def __init__(self, master, **kw):
        if not ClosableNotebook._style_done:
            ClosableNotebook._init_close_style(master)
        kw.setdefault("style", "Closable.TNotebook")
        super().__init__(master, **kw)
        self._press_idx   = None
        self.close_callback = None
        self.bind("<ButtonPress-1>",   self._on_press,   True)
        self.bind("<ButtonRelease-1>", self._on_release)

    @classmethod
    def _init_close_style(cls, master):
        style = ttk.Style(master)
        if HAS_PIL:
            sz = 16
            for state, color in [("n", "#999999"), ("a", "#cc3300")]:
                img = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
                d   = ImageDraw.Draw(img)
                d.line([(4, 4), (sz-5, sz-5)], fill=color, width=2)
                d.line([(sz-5, 4), (4, sz-5)], fill=color, width=2)
                photo = ImageTk.PhotoImage(img)
                if state == "n":
                    cls._img_n = photo
                else:
                    cls._img_a = photo
            style.element_create("Closable.close", "image",
                                 cls._img_n, ("active", cls._img_a),
                                 border=4, sticky="")
            style.layout("Closable.TNotebook.Tab", [
                ("Notebook.tab", {"sticky": "nswe", "children": [
                    ("Notebook.padding", {"side": "top", "sticky": "nswe", "children": [
                        ("Notebook.focus", {"side": "top", "sticky": "nswe", "children": [
                            ("Notebook.label",  {"side": "left", "sticky": ""}),
                            ("Closable.close",  {"side": "left", "sticky": ""}),
                        ]}),
                    ]}),
                ]}),
            ])
        else:
            style.layout("Closable.TNotebook.Tab", style.layout("TNotebook.Tab"))
        cls._style_done = True

    def _on_press(self, event):
        try:
            if self.identify(event.x, event.y) == "Closable.close":
                self._press_idx = self.index(f"@{event.x},{event.y}")
                return "break"
        except Exception:
            pass
        self._press_idx = None

    def _on_release(self, event):
        if self._press_idx is not None:
            try:
                if self.identify(event.x, event.y) == "Closable.close":
                    idx = self.index(f"@{event.x},{event.y}")
                    if idx == self._press_idx and callable(self.close_callback):
                        self.close_callback(idx=self._press_idx)
            except Exception:
                pass
            self._press_idx = None


# ================================================================ APPLICATION

class CourrierApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DacTime")
        self.geometry("780x740")
        self.minsize(640, 560)

        self._theme_nom   = "clair"
        self.T            = THEMES["clair"]
        self._lt_url      = None
        self._search_win  = None
        self._word_wrap   = tk.BooleanVar(value=True)
        self._statusbar_v = tk.BooleanVar(value=True)
        self._font_family = tk.StringVar(value="Courier New")
        self._font_size   = tk.IntVar(value=13)
        self._encodage    = tk.StringVar(value="utf-8")

        # Charger la config expéditeur (par utilisateur Windows)
        cfg = charger_config()
        self._exp_defaut = [
            cfg.get("exp_nom",     EXPEDITEUR[0]),
            cfg.get("exp_adresse", EXPEDITEUR[1]),
            cfg.get("exp_cpville", EXPEDITEUR[2]),
        ]
        self._tel_defaut = cfg.get("telephone", TELEPHONE)
        self._hist_exp   = cfg.get("historique_exp", [])
        global OUTPUT_DIR
        OUTPUT_DIR = cfg.get("output_dir", OUTPUT_DIR)

        self.configure(bg=self.T["bg"])
        self._build_menu()
        self._build_notebook()
        self._build_statusbar()
        self._bind_keys()
        self._set_icon()
        self._detect_lt()
        self.nouveau_onglet()

        if HAS_WINDND:
            windnd.hook_dropfiles(self, func=self._on_drop)

        self.protocol("WM_DELETE_WINDOW", self._quitter)

    # -------------------------------------------------------- icône plume

    def _set_icon(self):
        if not HAS_PIL:
            return
        try:
            img = self._draw_typewriter(64)
            self._icon = ImageTk.PhotoImage(img)
            self.iconphoto(True, self._icon)
        except Exception:
            pass

    @staticmethod
    def _draw_typewriter(sz=64):
        """Machine à écrire vintage 1950s (style Olivetti)."""
        img = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
        d   = ImageDraw.Draw(img)

        # Palette 1950s
        teal    = "#3a7a7c"   # corps teal vintage
        teal_l  = "#4e9698"   # corps clair (reflet)
        outline = "#162828"   # contour foncé
        chrome  = "#b0a890"   # chrome/métal
        chrome_l= "#d8d0b8"   # reflet chrome
        cream   = "#f0ead8"   # papier
        paper_l = "#fbf8f0"   # papier clair
        key_dk  = "#1e1e1e"   # touche foncée
        key_lt  = "#d8d0b0"   # face touche ivoire
        red     = "#c03030"   # plaque marque
        shadow  = (0, 0, 0, 60)

        s = sz / 64  # facteur d'échelle

        def sc(v):
            return int(v * s)

        # --- Ombre portée ---
        d.ellipse([sc(8), sc(58), sc(56), sc(64)], fill=shadow)

        # --- Corps principal ---
        d.rounded_rectangle([sc(3), sc(30), sc(61), sc(60)],
                             radius=sc(6), fill=teal, outline=outline, width=sc(1.5))
        # Reflet haut du corps
        d.rounded_rectangle([sc(5), sc(31), sc(59), sc(40)],
                             radius=sc(4), fill=teal_l, outline=None)

        # --- Platen (rouleau) ---
        d.rounded_rectangle([sc(5), sc(20), sc(59), sc(31)],
                             radius=sc(3), fill=chrome, outline=outline, width=sc(1.5))
        d.line([sc(7), sc(22), sc(57), sc(22)], fill=chrome_l, width=sc(1.5))

        # --- Papier ---
        d.rectangle([sc(17), sc(4), sc(47), sc(22)],
                    fill=cream, outline="#b0a880", width=sc(1))
        d.rectangle([sc(18), sc(5), sc(46), sc(9)],
                    fill=paper_l)
        for py in [10, 13, 16, 19]:
            d.line([sc(20), sc(py), sc(44), sc(py)], fill="#c8c0a0", width=sc(1))

        # --- Guides papier ---
        for gx in [sc(8), sc(51)]:
            d.rounded_rectangle([gx, sc(17), gx + sc(6), sc(26)],
                                 radius=sc(2), fill=chrome, outline=outline, width=sc(1))

        # --- Levier de retour chariot ---
        d.line([sc(56), sc(25), sc(62), sc(18)], fill=outline, width=sc(2))
        d.ellipse([sc(59), sc(16), sc(63), sc(21)], fill=chrome, outline=outline)

        # --- Clavier : 3 rangées ---
        rows = [
            (sc(34), 6, sc(8)),
            (sc(41), 7, sc(5)),
            (sc(48), 6, sc(8)),
        ]
        kw, kh = sc(7), sc(5)
        for (ky, nk, kx_start) in rows:
            for ki in range(nk):
                kx = kx_start + ki * sc(8)
                d.rounded_rectangle([kx, ky, kx + kw, ky + kh],
                                     radius=sc(1.5), fill=key_dk, outline=outline, width=sc(1))
                d.rounded_rectangle([kx + sc(1), ky + sc(1), kx + kw - sc(1), ky + kh - sc(1)],
                                     radius=sc(1), fill=key_lt)

        # --- Barre d'espace ---
        d.rounded_rectangle([sc(10), sc(55), sc(54), sc(59)],
                             radius=sc(2), fill=key_dk, outline=outline, width=sc(1))

        # --- Plaque de marque (rouge, centré sur le corps) ---
        d.rounded_rectangle([sc(22), sc(33), sc(42), sc(37)],
                             radius=sc(1.5), fill=red, outline="#801818")

        return img

    # -------------------------------------------------------- menu

    def _build_menu(self):
        T  = self.T
        mb = tk.Menu(self, bg=T["bg2"], fg=T["fg"],
                     activebackground="#2a6099", activeforeground="white")
        self.configure(menu=mb)

        def m(parent, **kw):
            return tk.Menu(parent, tearoff=0, bg=T["bg2"], fg=T["fg"],
                           activebackground="#2a6099", activeforeground="white", **kw)

        # Fichier
        mf = m(mb); mb.add_cascade(label="Fichier", menu=mf)
        mf.add_command(label="Nouveau courrier   Ctrl+N",  command=self._nouveau_courrier)
        mf.add_command(label="Nouveau onglet     Ctrl+T",  command=self.nouveau_onglet)
        mf.add_command(label="Fermer onglet      Ctrl+W",  command=self.fermer_onglet)
        mf.add_separator()
        mf.add_command(label="Ouvrir...          Ctrl+O",  command=self._ouvrir_fichier)
        mf.add_separator()
        mf.add_command(label="Enregistrer        Ctrl+S",  command=self._enregistrer_rapide)
        mf.add_command(label="Enregistrer sous...",        command=self._enregistrer)
        mf.add_command(label="Exporter en PDF...",         command=self._exporter_pdf)
        mf.add_separator()
        mf.add_command(label="Imprimer           Ctrl+P",  command=self._imprimer)
        mf.add_command(label="Imprimer enveloppe Ctrl+E",  command=self._imprimer_enveloppe)
        mf.add_command(label="Ouvrir dans Bloc-notes",     command=self._ouvrir_blocnotes)
        mf.add_separator()
        mf.add_command(label="Quitter",                    command=self._quitter)

        # Édition
        me = m(mb); mb.add_cascade(label="Édition", menu=me)
        me.add_command(label="Annuler            Ctrl+Z",
                       command=lambda: self._tab().corps.event_generate("<<Undo>>"))
        me.add_command(label="Rétablir           Ctrl+Y",
                       command=lambda: self._tab().corps.event_generate("<<Redo>>"))
        me.add_separator()
        me.add_command(label="Couper             Ctrl+X",
                       command=lambda: self._txt_actif().event_generate("<<Cut>>"))
        me.add_command(label="Copier             Ctrl+C",
                       command=lambda: self._txt_actif().event_generate("<<Copy>>"))
        me.add_command(label="Coller             Ctrl+V",
                       command=lambda: self._txt_actif().event_generate("<<Paste>>"))
        me.add_separator()
        me.add_command(label="Rechercher...      Ctrl+F",  command=self._rechercher)
        me.add_command(label="Suivant            F3",      command=self._rechercher_suivant)
        me.add_command(label="Précédent          Maj+F3",  command=self._rechercher_precedent)
        me.add_command(label="Remplacer...       Ctrl+H",  command=self._remplacer)
        me.add_command(label="Atteindre...       Ctrl+G",  command=self._aller_a_ligne)
        me.add_separator()
        me.add_command(label="Sélectionner tout  Ctrl+A",
                       command=lambda: self._tab().corps.tag_add(tk.SEL, "1.0", "end"))
        me.add_command(label="Heure et date      F5",      command=self._inserer_date)
        me.add_separator()
        me.add_command(label="Copier cette page",          command=self._copier_page)
        me.add_command(label="Copier toutes les pages",    command=self._copier_tout)

        # Modèles
        mm = m(mb); mb.add_cascade(label="Modèles", menu=mm)
        for nom_modele in MODELES:
            mm.add_command(label=nom_modele,
                           command=lambda n=nom_modele: self._charger_modele(n))

        # Destinataires récents
        self._menu_hist = m(mb)
        mb.add_cascade(label="Récents", menu=self._menu_hist)
        self._rafraichir_menu_hist()

        # Format
        mfmt = m(mb); mb.add_cascade(label="Format", menu=mfmt)
        mfmt.add_checkbutton(label="Retour à la ligne",
                             variable=self._word_wrap, command=self._toggle_wrap)
        mfmt.add_command(label="Police...",  command=self._choisir_police)
        mfmt.add_separator()
        m_enc = m(mfmt); mfmt.add_cascade(label="Encodage", menu=m_enc)
        for enc, lbl in [("utf-8", "UTF-8"), ("utf-8-sig", "UTF-8 avec BOM"),
                         ("cp1252", "ANSI (Windows-1252)"), ("utf-16", "UTF-16")]:
            m_enc.add_radiobutton(label=lbl, variable=self._encodage, value=enc)

        # Affichage
        ma = m(mb); mb.add_cascade(label="Affichage", menu=ma)
        ma.add_command(label="Thème sombre / clair",       command=self._basculer_theme)
        ma.add_separator()
        mz = m(ma); ma.add_cascade(label="Zoom", menu=mz)
        mz.add_command(label="Zoom avant   Ctrl++", command=self._zoom_plus)
        mz.add_command(label="Zoom arrière Ctrl+-", command=self._zoom_moins)
        mz.add_command(label="Normal       Ctrl+0", command=self._zoom_reset)
        ma.add_separator()
        ma.add_checkbutton(label="Barre d'état",
                           variable=self._statusbar_v, command=self._toggle_statusbar)

        # Correcteur
        mc = m(mb); mb.add_cascade(label="Correcteur", menu=mc)
        mc.add_command(label="Vérifier l'orthographe  F7", command=self._verifier)

        # Paramètres
        mp = m(mb); mb.add_cascade(label="Paramètres", menu=mp)
        mp.add_command(label="Expéditeur par défaut...", command=self._config_expediteur)
        mp.add_command(label="Dossier de sauvegarde...", command=self._choisir_dossier)
        mp.add_separator()
        mp.add_command(label="Thème sombre / clair",    command=self._basculer_theme)

        # Aide
        mh = m(mb); mb.add_cascade(label="Aide", menu=mh)
        mh.add_command(label="Raccourcis clavier...   F1",    command=self._raccourcis)
        mh.add_command(label="Notes de version...",           command=self._notes_version)
        mh.add_separator()
        mh.add_command(label="À propos de DacTime...",        command=self._a_propos)

    # -------------------------------------------------------- notebook

    def _build_notebook(self):
        T = self.T
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook",
                        background=T["bg"], borderwidth=0)
        style.configure("TNotebook.Tab",
                        background=T["bg2"], foreground=T["fg"],
                        padding=[12, 4], font=("Courier New", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", T["bg3"]), ("active", T["bg3"])],
                  foreground=[("selected", T["fg"]),  ("active", T["fg"])])

        self.notebook = ClosableNotebook(self, style="Closable.TNotebook")
        self.notebook.close_callback = self.fermer_onglet
        self.notebook.pack(fill="both", expand=True, padx=0, pady=0)
        self.notebook.bind("<<NotebookTabChanged>>", lambda e: self._update_statusbar())
        self.notebook.bind("<Button-3>", self._tab_right_click)

    def _build_statusbar(self):
        T = self.T
        self._statusbar = tk.Frame(self, bg=T["bg2"], height=22)
        self._statusbar.pack(fill="x", side="bottom")
        self._statut    = tk.StringVar(value="")
        tk.Label(self._statusbar, textvariable=self._statut,
                 font=("Courier New", 9), bg=T["bg2"], fg=T["fg3"],
                 anchor="w", padx=10).pack(side="left")
        self._lbl_pos   = tk.Label(self._statusbar, text="Ln 1, Col 1",
                                   font=("Courier New", 9), bg=T["bg2"], fg=T["fg3"], padx=10)
        self._lbl_pos.pack(side="right")
        self._lbl_words = tk.Label(self._statusbar, text="0 mot(s)",
                                   font=("Courier New", 9), bg=T["bg2"], fg=T["fg3"], padx=10)
        self._lbl_words.pack(side="right")
        self._lbl_zoom  = tk.Label(self._statusbar, text="100%",
                                   font=("Courier New", 9), bg=T["bg2"], fg=T["fg3"], padx=10)
        self._lbl_zoom.pack(side="right")

    def _update_statusbar(self):
        try:
            tab = self._tab()
            focused = self.focus_get()
            txt_actif = tab.corps
            for t in tab._corps_pages:
                if t == focused:
                    txt_actif = t
                    break
            nb_p   = getattr(txt_actif, '_page_num',   1)
            max_l  = getattr(txt_actif, '_max_lignes', 38)
            nb_l   = _lignes_impression(txt_actif)
            reste  = max_l - nb_l
            if nb_l >= max_l:
                label = f"P{nb_p} PLEINE — + Nouvelle page"
            elif reste <= 5:
                label = f"P{nb_p} : {nb_l}/{max_l} lignes  ({reste} restante(s))"
            else:
                label = f"P{nb_p} : {nb_l}/{max_l} lignes"
            self._lbl_pos.config(text=label)
            texte = txt_actif.get("1.0", "end-1c")
            nb    = len(texte.split()) if texte.strip() else 0
            self._lbl_words.config(text=f"{nb} mot(s)")
        except Exception:
            pass

    def _update_tab_title(self, tab):
        try:
            self.notebook.tab(tab, text=tab.get_titre())
        except Exception:
            pass

    # -------------------------------------------------------- onglets

    def nouveau_onglet(self):
        tab = CourrierTab(self.notebook, self)
        self.notebook.add(tab, text="Nouveau courrier")
        self.notebook.select(tab)
        tab.corps.focus_set()

    def fermer_onglet(self, idx=None):
        tabs = self.notebook.tabs()
        if len(tabs) <= 1:
            messagebox.showinfo("Fermer", "Il doit rester au moins un onglet.")
            return
        tab = self.notebook.nametowidget(tabs[idx]) if idx is not None else self._tab()
        if self._onglet_modifie(tab):
            rep = messagebox.askyesnocancel(
                "Fermer l'onglet",
                f"Le courrier « {tab.get_titre()} » contient du contenu.\n"
                "Enregistrer avant de fermer ?")
            if rep is None:
                return
            if rep:
                self._sauvegarder_temp(tab)
        self.notebook.forget(tab)

    def _onglet_modifie(self, tab):
        corps_vide = all(not t.get("1.0", "end-1c").strip() for t in tab._corps_pages)
        return bool(not corps_vide or tab.nom.get().strip() or tab.d_objet.get().strip())

    def _tab_right_click(self, event):
        try:
            self.notebook.select(self.notebook.index(f"@{event.x},{event.y}"))
        except Exception:
            pass
        T    = self.T
        menu = tk.Menu(self, tearoff=0, bg=T["bg2"], fg=T["fg"],
                       activebackground="#2a6099", activeforeground="white")
        menu.add_command(label="Nouvel onglet    Ctrl+T", command=self.nouveau_onglet)
        menu.add_command(label="Fermer l'onglet  Ctrl+W", command=self.fermer_onglet)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _tab(self):
        return self.notebook.nametowidget(self.notebook.select())

    # -------------------------------------------------------- thème

    def _basculer_theme(self):
        self._theme_nom = "sombre" if self._theme_nom == "clair" else "clair"
        self.T = THEMES[self._theme_nom]
        self._appliquer_theme()

    def _appliquer_theme(self):
        T = self.T
        self.configure(bg=T["bg"])

        # Style Notebook
        style = ttk.Style(self)
        style.configure("TNotebook", background=T["bg"])
        style.configure("TNotebook.Tab", background=T["bg2"], foreground=T["fg"])
        style.map("TNotebook.Tab",
                  background=[("selected", T["bg3"]), ("active", T["bg3"])],
                  foreground=[("selected", T["fg"]),  ("active", T["fg"])])

        # Barre d'état
        self._statusbar.configure(bg=T["bg2"])
        for w in self._statusbar.winfo_children():
            w.configure(bg=T["bg2"], fg=T["fg3"])

        # Onglets existants — recréer
        for tab_id in self.notebook.tabs():
            old_tab = self.notebook.nametowidget(tab_id)
            titre   = old_tab.get_titre()
            contenu = {
                "exp_nom":     old_tab.exp_nom.get(),
                "exp_adresse": old_tab.exp_adresse.get(),
                "exp_cpville": old_tab.exp_cpville.get(),
                "nom":     old_tab.nom.get(),
                "societe": old_tab.societe.get(),
                "adresse": old_tab.adresse.get(),
                "cpville": old_tab.cpville.get(),
                "date":    old_tab.d_date.get(),
                "objet":   old_tab.d_objet.get(),
                "pages":   [t.get("1.0", "end-1c") for t in old_tab._corps_pages],
            }
            self.notebook.forget(tab_id)
            new_tab = CourrierTab(self.notebook, self)
            new_tab.exp_nom.set(contenu["exp_nom"])
            new_tab.exp_adresse.set(contenu["exp_adresse"])
            new_tab.exp_cpville.set(contenu["exp_cpville"])
            new_tab.nom.set(contenu["nom"])
            new_tab.societe.set(contenu["societe"])
            new_tab.adresse.set(contenu["adresse"])
            new_tab.cpville.set(contenu["cpville"])
            new_tab.d_date.set(contenu["date"])
            new_tab.d_objet.set(contenu["objet"])
            new_tab.corps.insert("1.0", contenu["pages"][0] if contenu["pages"] else "")
            for pg in contenu["pages"][1:]:
                new_tab._ajouter_page()
                new_tab._corps_pages[-1].insert("1.0", pg)
            self.notebook.add(new_tab, text=titre or "Nouveau courrier")

    # -------------------------------------------------------- raccourcis

    def _bind_keys(self):
        self.bind_all("<F1>",            lambda e: self._raccourcis())
        self.bind_all("<Control-n>",     lambda e: self._nouveau_courrier())
        self.bind_all("<Control-t>",     lambda e: self.nouveau_onglet())
        self.bind_all("<Control-w>",     lambda e: self.fermer_onglet())
        self.bind_all("<Control-o>",     lambda e: self._ouvrir_fichier())
        self.bind_all("<Control-s>",     lambda e: self._enregistrer_rapide())
        self.bind_all("<Control-S>",     lambda e: self._enregistrer())
        self.bind_all("<Control-p>",     lambda e: self._imprimer())
        self.bind_all("<Control-e>",     lambda e: self._imprimer_enveloppe())
        self.bind_all("<Control-f>",     lambda e: self._rechercher())
        self.bind_all("<Control-h>",     lambda e: self._remplacer())
        self.bind_all("<Control-g>",     lambda e: self._aller_a_ligne())
        self.bind_all("<Control-a>",     lambda e: self._tab().corps.tag_add(tk.SEL, "1.0", "end"))
        self.bind_all("<F3>",            lambda e: self._rechercher_suivant())
        self.bind_all("<Shift-F3>",      lambda e: self._rechercher_precedent())
        self.bind_all("<F5>",            lambda e: self._inserer_date())
        self.bind_all("<F7>",            lambda e: self._verifier())
        self.bind_all("<Control-equal>", lambda e: self._zoom_plus())
        self.bind_all("<Control-minus>", lambda e: self._zoom_moins())
        self.bind_all("<Control-0>",     lambda e: self._zoom_reset())

    # -------------------------------------------------------- actions fichier

    def _sauvegarder_temp(self, tab=None):
        tab = tab or self._tab()
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        path = os.path.join(OUTPUT_DIR, tab.nom_fichier())
        with open(path, "w", encoding="utf-8") as f:
            json.dump(tab.build_dactime(), f, ensure_ascii=False, indent=2)
        dest = {"nom": tab.nom.get(), "societe": tab.societe.get(),
                "adresse": tab.adresse.get(), "cpville": tab.cpville.get()}
        if dest["nom"]:
            ajouter_historique(dest)
            self._rafraichir_menu_hist()
        self._sauvegarder_expediteur_cache(tab)
        return path

    def _enregistrer_rapide(self):
        tab = self._tab()
        path = self._sauvegarder_temp(tab)
        self._statut.set(f"Enregistré : {path}")

    def _enregistrer(self):
        tab = self._tab()
        path = filedialog.asksaveasfilename(
            initialdir=OUTPUT_DIR, initialfile=tab.nom_fichier(),
            defaultextension=".dactime",
            filetypes=[("Courrier DacTime", "*.dactime"), ("Fichier texte", "*.txt")],
            title="Enregistrer",
        )
        if path:
            if path.lower().endswith(".dactime"):
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(tab.build_dactime(), f, ensure_ascii=False, indent=2)
            else:
                with open(path, "w", encoding=self._encodage.get()) as f:
                    f.write(tab.build_txt())
            self._statut.set(f"Enregistré : {path}")

    def _txt_temp(self, tab):
        import tempfile
        # Supprime le séparateur de page (\f entouré de \n) pour que la page 2
        # commence à la ligne 1 de la feuille 2 sans ligne vide parasite.
        contenu = tab.build_txt().replace("\n\f\n", "\n")
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt",
                                         encoding="utf-8", delete=False)
        tmp.write(contenu)
        tmp.close()
        return tmp.name

    def _ouvrir_blocnotes(self):
        tab = self._tab()
        subprocess.Popen(["notepad.exe", self._txt_temp(tab)])

    def _imprimer(self):
        tab = self._tab()
        path = self._txt_temp(tab)
        import tempfile
        ps = tempfile.NamedTemporaryFile(mode="w", suffix=".ps1",
                                         delete=False, encoding="utf-8")
        ps.write(
            f'$path = "{path}"\n'
            f'$p = Start-Process notepad.exe -ArgumentList $path -PassThru\n'
            f'Start-Sleep -Milliseconds 2000\n'
            f'$w = New-Object -ComObject WScript.Shell\n'
            f'$w.AppActivate($p.Id)\n'
            f'Start-Sleep -Milliseconds 300\n'
            f"$w.SendKeys('^p')\n"
        )
        ps.close()
        subprocess.Popen(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", ps.name],
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        self._statut.set(f"Impression : {tab.nom_fichier()}")

    def _imprimer_enveloppe(self):
        tab = self._tab()
        exp  = [tab.exp_nom.get().strip(),
                tab.exp_adresse.get().strip(),
                tab.exp_cpville.get().strip()]
        dest = [tab.nom.get().strip(),
                tab.societe.get().strip(),
                tab.adresse.get().strip(),
                tab.cpville.get().strip()]
        exp  = [l for l in exp  if l]
        dest = [l for l in dest if l]
        if not dest:
            messagebox.showwarning("Enveloppe", "Aucun destinataire renseigné.")
            return
        if HAS_PDF:
            self._enveloppe_pdf(exp, dest)
        else:
            messagebox.showerror("Enveloppe",
                "ReportLab non disponible.\npip install reportlab")

    def _enveloppe_win32(self, exp, dest):
        try:
            # CreatePrintDialog(bPrintSetupOnly, dwFlags, parent)
            # PD_RETURNDC (0x100) : le dialogue alloue et retourne un DC utilisable
            PD_RETURNDC = getattr(win32con, "PD_RETURNDC", 0x100)  # type: ignore[name-defined]
            dlg = win32ui.CreatePrintDialog(False, PD_RETURNDC, None)  # type: ignore[name-defined]
            if dlg.DoModal() != 1:  # 1 = IDOK
                return  # annule par l'utilisateur

            # DC de l'imprimante telle que configuree dans le dialogue
            hdc = dlg.GetPrinterDC()
            if hdc is None:
                raise RuntimeError("GetPrinterDC() returned None")
            dc = win32ui.CreateDCFromHandle(hdc)  # type: ignore[name-defined]
            dc.SetMapMode(win32con.MM_LOMETRIC)  # type: ignore[name-defined]
            # MM_LOMETRIC : 1 unite = 0.1mm, Y negatif vers le bas
            # DL paysage : 2200 unites large x 1100 unites haut

            dc.StartDoc("DacTime - Enveloppe DL")
            dc.StartPage()

            # Police expediteur (petit, ~8pt)
            f_exp = win32ui.CreateFont({  # type: ignore[name-defined]
                "name": "Courier New", "height": -28, "weight": 400})
            dc.SelectObject(f_exp)
            y = -80
            for ligne in exp:
                dc.TextOut(80, y, ligne)
                y -= 45

            # Police destinataire (~11pt gras)
            f_dest = win32ui.CreateFont({  # type: ignore[name-defined]
                "name": "Courier New", "height": -39, "weight": 700})
            dc.SelectObject(f_dest)
            interligne = 60
            y_dest = -(550 - (len(dest) - 1) * interligne // 2)
            for ligne in dest:
                dc.TextOut(1100, y_dest, ligne)
                y_dest -= interligne

            dc.EndPage()
            dc.EndDoc()
            dc.DeleteDC()
            self._statut.set("Enveloppe DL envoyee a l'imprimante")

        except Exception as e:
            messagebox.showerror("Enveloppe",
                f"Erreur impression :\n{e}\n\nBascule sur PDF.")
            if HAS_PDF:
                self._enveloppe_pdf(exp, dest)

    def _enveloppe_pdf(self, exp, dest):
        import tempfile
        mm = cm / 10  # type: ignore[name-defined]
        DL_W, DL_H = 220 * mm, 110 * mm
        path = tempfile.mktemp(suffix=".pdf")
        c = rl_canvas.Canvas(path, pagesize=(DL_W, DL_H))  # type: ignore[name-defined]
        c.setFont("Courier", 8)
        y = DL_H - 12 * mm
        for ligne in exp:
            c.drawString(8 * mm, y, ligne)
            y -= 10
        c.setFont("Courier-Bold", 11)
        interligne = 15
        y_dest = min(DL_H / 2 + (len(dest) * interligne / 2), DL_H - 40 * mm)
        for ligne in dest:
            c.drawString(110 * mm, y_dest, ligne)
            y_dest -= interligne
        c.save()
        os.startfile(path)
        self._statut.set("Enveloppe DL prete - imprimer depuis le lecteur PDF")

    def _build_pdf(self, tab, path):
        if not HAS_PDF:
            return
        txt  = tab.build_txt()
        c    = rl_canvas.Canvas(path, pagesize=A4)  # type: ignore[name-defined]
        w, h = A4                                    # type: ignore[name-defined]
        marge_g, marge_h = 2.5 * cm, 2.5 * cm      # type: ignore[name-defined]
        c.setFont("Courier", 10)
        x, y  = marge_g, h - marge_h
        pas_y = 14
        for ligne in txt.split("\n"):
            if ligne == "\f":
                c.showPage(); c.setFont("Courier", 10); y = h - marge_h
                continue
            ligne_affichee = ligne.replace("\t", " " * 8)
            if y < marge_h + pas_y:
                c.showPage(); c.setFont("Courier", 10); y = h - marge_h
            c.drawString(x, y, ligne_affichee)
            y -= pas_y
        c.save()

    def _ouvrir_fichier(self):
        path = filedialog.askopenfilename(
            initialdir=OUTPUT_DIR,
            filetypes=[("Courrier DacTime", "*.dactime"), ("Fichier texte", "*.txt"), ("Tous", "*.*")],
            title="Ouvrir",
        )
        if path:
            self._charger_fichier(path)

    def _charger_fichier(self, path):
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                contenu = f.read()
            tab = self._tab()
            if path.lower().endswith(".dactime"):
                data = json.loads(contenu)
                tab.exp_nom.set(data.get("exp_nom", ""))
                tab.exp_adresse.set(data.get("exp_adresse", ""))
                tab.exp_cpville.set(data.get("exp_cpville", ""))
                tab.nom.set(data.get("dest_nom", ""))
                tab.societe.set(data.get("dest_societe", ""))
                tab.adresse.set(data.get("dest_adresse", ""))
                tab.cpville.set(data.get("dest_cpville", ""))
                tab.d_date.set(data.get("date", date.today().strftime("%d/%m/%Y")))
                tab.d_objet.set(data.get("objet", ""))
                pages = data.get("pages", [data.get("corps", "")])
                tab.corps.delete("1.0", "end")
                tab.corps.insert("1.0", pages[0] if pages else "")
                for pg in pages[1:]:
                    tab._ajouter_page()
                    tab._corps_pages[-1].delete("1.0", "end")
                    tab._corps_pages[-1].insert("1.0", pg)
                self._update_tab_title(tab)
            else:
                tab.corps.delete("1.0", "end")
                tab.corps.insert("1.0", contenu)
            tab._effacer_erreurs()
            self._statut.set(f"Ouvert : {path}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ouvrir :\n{e}")

    def _on_drop(self, files):
        if files:
            path = files[0]
            if isinstance(path, bytes):
                path = path.decode("cp1252", errors="replace")
            self._charger_fichier(path)

    def _exporter_pdf(self):
        tab = self._tab()
        if not HAS_PDF:
            messagebox.showerror("PDF", "reportlab non installé.\nLancer : pip install reportlab")
            return
        path = filedialog.asksaveasfilename(
            initialdir=OUTPUT_DIR,
            initialfile=tab.nom_fichier().replace(".dactime", ".pdf"),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            title="Exporter en PDF",
        )
        if not path:
            return

        self._build_pdf(tab, path)
        self._statut.set(f"PDF exporté : {path}")
        os.startfile(path)

    def _sauvegarder_expediteur_cache(self, tab=None):
        """Ajoute l'expéditeur courant à l'historique (5 max). Ne touche pas à l'expéditeur par défaut."""
        if tab is None:
            try:
                tab = self._tab()
            except Exception:
                return
        nom = tab.exp_nom.get().strip()
        if not nom:
            return
        entree = {
            "exp_nom":     nom,
            "exp_adresse": tab.exp_adresse.get().strip(),
            "exp_cpville": tab.exp_cpville.get().strip(),
            "telephone":   self._tel_defaut,
        }
        cfg  = charger_config()
        hist = cfg.get("historique_exp", [])
        hist = [h for h in hist if h.get("exp_nom") != nom]
        hist.insert(0, entree)
        cfg["historique_exp"] = hist[:5]
        sauvegarder_config(cfg)
        self._hist_exp = cfg["historique_exp"]

    def _popup_smart(self, pop):
        """Positionne pop près du curseur, vers le centre de la fenêtre, dans ses limites."""
        mx = self.winfo_pointerx()
        my = self.winfo_pointery()
        cx = self.winfo_rootx() + self.winfo_width()  // 2
        cy = self.winfo_rooty() + self.winfo_height() // 2
        def _do():
            pop.update_idletasks()
            pw, ph = pop.winfo_reqwidth(), pop.winfo_reqheight()
            ax = self.winfo_rootx()
            ay = self.winfo_rooty()
            aw = self.winfo_width()
            ah = self.winfo_height()
            px = (mx - pw - 8) if mx >= cx else (mx + 8)
            py = (my - ph - 8) if my >= cy else (my + 8)
            px = max(ax, min(px, ax + aw - pw))
            py = max(ay, min(py, ay + ah - ph))
            pop.geometry(f"+{px}+{py}")
        pop.after(1, _do)

    def _choisir_dossier(self):
        global OUTPUT_DIR
        dossier = filedialog.askdirectory(
            initialdir=OUTPUT_DIR,
            title="Dossier de sauvegarde par défaut",
        )
        if dossier:
            OUTPUT_DIR = dossier
            cfg = charger_config()
            cfg["output_dir"] = dossier
            sauvegarder_config(cfg)
            self._statut.set(f"Dossier : {dossier}")

    def _raccourcis(self):
        T   = self.T
        pop = tk.Toplevel(self)
        pop.title("Raccourcis clavier")
        pop.resizable(False, False)
        pop.configure(bg=T["bg"])
        pop.grab_set()
        pop.focus_set()

        tk.Label(pop, text="Raccourcis clavier", font=("Courier New", 14, "bold"),
                 bg=T["bg"], fg=T["fg"]).pack(pady=(20, 4))
        tk.Frame(pop, bg=T["sep"], height=1).pack(fill="x", padx=24, pady=10)

        raccourcis = [
            ("FICHIER", None),
            ("Ctrl+N",     "Nouveau courrier"),
            ("Ctrl+T",     "Nouvel onglet"),
            ("Ctrl+W",     "Fermer l'onglet"),
            ("Ctrl+O",     "Ouvrir un fichier"),
            ("Ctrl+S",     "Enregistrer"),
            ("Ctrl+Maj+S", "Enregistrer sous..."),
            ("Ctrl+P",     "Imprimer le courrier"),
            ("Ctrl+E",     "Imprimer une enveloppe DL"),
            ("EDITION", None),
            ("Ctrl+Z",     "Annuler"),
            ("Ctrl+Y",     "Retablir"),
            ("Ctrl+X / C / V", "Couper / Copier / Coller"),
            ("Ctrl+A",     "Selectionner tout"),
            ("Ctrl+F",     "Rechercher"),
            ("Ctrl+H",     "Rechercher et remplacer"),
            ("F3 / Maj+F3","Occurrence suivante / precedente"),
            ("Ctrl+G",     "Aller a la ligne"),
            ("F5",         "Inserer la date du jour"),
            ("AFFICHAGE", None),
            ("Ctrl++",     "Zoom avant"),
            ("Ctrl+-",     "Zoom arriere"),
            ("Ctrl+0",     "Zoom normal"),
            ("OUTILS", None),
            ("F7",         "Verifier l'orthographe"),
            ("AIDE", None),
            ("F1",         "Raccourcis clavier (cette fenetre)"),
        ]

        frm = tk.Frame(pop, bg=T["bg"])
        frm.pack(padx=28, pady=4)
        row = 0
        for item in raccourcis:
            key, desc = item
            if desc is None:
                tk.Label(frm, text=key, font=("Courier New", 9, "bold"),
                         bg=T["bg"], fg=T["fg2"], anchor="w").grid(
                    row=row, column=0, columnspan=2, sticky="w", pady=(10, 2))
            else:
                tk.Label(frm, text=key, font=("Courier New", 10, "bold"),
                         bg=T["bg"], fg=T["fg"], anchor="w", width=18).grid(
                    row=row, column=0, sticky="w")
                tk.Label(frm, text=desc, font=("Courier New", 10),
                         bg=T["bg"], fg=T["fg2"], anchor="w").grid(
                    row=row, column=1, sticky="w", padx=(8, 0))
            row += 1

        tk.Frame(pop, bg=T["sep"], height=1).pack(fill="x", padx=24, pady=10)
        tk.Button(pop, text="Fermer", font=("Courier New", 9),
                  bg=T["bg3"], fg=T["fg"], relief="flat", bd=0,
                  cursor="hand2", padx=20, pady=6,
                  command=pop.destroy).pack(pady=(0, 20))
        pop.bind("<Escape>", lambda e: pop.destroy())
        pop.bind("<F1>",     lambda e: pop.destroy())
        self._popup_smart(pop)

    def _notes_version(self):
        T   = self.T
        pop = tk.Toplevel(self)
        pop.title("Notes de version")
        pop.resizable(False, False)
        pop.configure(bg=T["bg"])
        pop.grab_set()
        pop.focus_set()

        tk.Label(pop, text="Notes de version", font=("Courier New", 14, "bold"),
                 bg=T["bg"], fg=T["fg"]).pack(pady=(20, 4))
        tk.Frame(pop, bg=T["sep"], height=1).pack(fill="x", padx=24, pady=10)

        notes = [
            ("1.2", "juillet 2026", [
                "Impression d'enveloppes DL (Ctrl+E)",
                "Raccourcis clavier et notes de version (Aide)",
                "Onglets fermables avec bouton x",
                "Calendrier integre (clic droit sur date)",
                "Correcteur orthographique LanguageTool (F7)",
                "Zoom texte (Ctrl++ / Ctrl+-)",
                "Themes clair / sombre",
                "Rechercher et remplacer (Ctrl+H)",
                "Sauvegarde automatique format .dactime",
                "Association fichier .dactime",
            ]),
        ]

        frm = tk.Frame(pop, bg=T["bg"])
        frm.pack(padx=28, pady=4, fill="x")
        for version, date_ver, items in notes:
            tk.Label(frm, text=f"Version {version}  -  {date_ver}",
                     font=("Courier New", 11, "bold"),
                     bg=T["bg"], fg=T["fg"], anchor="w").pack(anchor="w", pady=(4, 2))
            for item in items:
                tk.Label(frm, text=f"  - {item}",
                         font=("Courier New", 9),
                         bg=T["bg"], fg=T["fg2"], anchor="w").pack(anchor="w")

        tk.Frame(pop, bg=T["sep"], height=1).pack(fill="x", padx=24, pady=10)
        tk.Button(pop, text="Fermer", font=("Courier New", 9),
                  bg=T["bg3"], fg=T["fg"], relief="flat", bd=0,
                  cursor="hand2", padx=20, pady=6,
                  command=pop.destroy).pack(pady=(0, 20))
        pop.bind("<Escape>", lambda e: pop.destroy())
        self._popup_smart(pop)

    def _a_propos(self):
        T   = self.T
        pop = tk.Toplevel(self)
        pop.title("À propos de DacTime")
        pop.resizable(False, False)
        pop.configure(bg=T["bg"])
        pop.grab_set()
        pop.focus_set()

        tk.Label(pop, text="DacTime", font=("Courier New", 22, "bold"),
                 bg=T["bg"], fg=T["fg"]).pack(pady=(28, 4))
        tk.Label(pop, text=f"Version {VERSION}",
                 font=("Courier New", 10), bg=T["bg"], fg=T["fg2"]).pack()

        tk.Frame(pop, bg=T["sep"], height=1).pack(fill="x", padx=30, pady=16)

        tk.Label(pop,
                 text="Application de rédaction\net d'impression de courriers",
                 font=("Courier New", 10), bg=T["bg"], fg=T["fg"],
                 justify="center").pack()

        tk.Frame(pop, bg=T["sep"], height=1).pack(fill="x", padx=30, pady=16)

        tk.Label(pop, text="Conçu et développé par",
                 font=("Courier New", 9), bg=T["bg"], fg=T["fg2"]).pack()
        tk.Label(pop, text="Clément LATTAR",
                 font=("Courier New", 13, "bold"), bg=T["bg"], fg=T["fg"]).pack(pady=(2, 0))
        tk.Label(pop, text="Technicien IT — SolidarIT",
                 font=("Courier New", 9, "italic"), bg=T["bg"], fg=T["fg2"]).pack()
        tk.Label(pop, text="mebalat@gmail.com",
                 font=("Courier New", 9), bg=T["bg"], fg=T["fg3"]).pack(pady=(2, 0))

        tk.Frame(pop, bg=T["sep"], height=1).pack(fill="x", padx=30, pady=16)

        tk.Label(pop, text=f"© {date.today().year} — Tous droits réservés",
                 font=("Courier New", 8), bg=T["bg"], fg=T["fg3"]).pack()

        tk.Button(pop, text="Fermer", font=("Courier New", 9),
                  bg=T["bg3"], fg=T["fg"], relief="flat", bd=0,
                  cursor="hand2", padx=20, pady=6,
                  command=pop.destroy).pack(pady=(16, 24))

        self._popup_smart(pop)

    def _config_expediteur(self):
        pop = tk.Toplevel(self)
        pop.title("Expediteur par defaut")
        pop.resizable(False, False)
        pop.grab_set()
        T = self.T
        pop.configure(bg=T["bg"])

        v_nom     = tk.StringVar(value=self._exp_defaut[0])
        v_adresse = tk.StringVar(value=self._exp_defaut[1])
        v_cpville = tk.StringVar(value=self._exp_defaut[2])
        v_tel     = tk.StringVar(value=self._tel_defaut)

        champs = [
            ("Nom / Prenom", v_nom),
            ("Adresse",      v_adresse),
            ("CP + Ville",   v_cpville),
            ("Telephone",    v_tel),
        ]
        for i, (label, var) in enumerate(champs):
            tk.Label(pop, text=label, font=("Courier New", 9), bg=T["bg"],
                     fg=T["fg2"], anchor="w", width=14).grid(
                         row=i, column=0, padx=(14, 4), pady=5, sticky="w")
            tk.Entry(pop, textvariable=var, font=("Courier New", 11), width=32,
                     bg=T["entry_bg"], fg=T["entry_fg"],
                     insertbackground=T["cursor"], relief="solid", bd=1).grid(
                         row=i, column=1, padx=(0, 14), pady=5, sticky="ew")

        tk.Label(pop, text="Ces valeurs seront pre-remplies a chaque ouverture.",
                 font=("Courier New", 8), bg=T["bg"], fg=T["fg3"]).grid(
                     row=len(champs), column=0, columnspan=2, pady=(2, 6))

        def enregistrer():
            cfg = charger_config()
            cfg["exp_nom"]     = v_nom.get().strip()
            cfg["exp_adresse"] = v_adresse.get().strip()
            cfg["exp_cpville"] = v_cpville.get().strip()
            cfg["telephone"]   = v_tel.get().strip()
            sauvegarder_config(cfg)
            self._exp_defaut = [cfg["exp_nom"], cfg["exp_adresse"], cfg["exp_cpville"]]
            self._tel_defaut = cfg["telephone"]
            pop.destroy()
            self._statut.set("Expediteur par defaut enregistre.")

        bf = tk.Frame(pop, bg=T["bg"])
        bf.grid(row=len(champs) + 1, column=0, columnspan=2, pady=(0, 12))
        tk.Button(bf, text="Enregistrer", font=("Courier New", 11), relief="flat",
                  bg="#2a6099", fg="white", padx=14, pady=6, cursor="hand2",
                  command=enregistrer).pack(side="left", padx=6)
        tk.Button(bf, text="Annuler", font=("Courier New", 9), relief="flat",
                  bg=T["bg2"], fg=T["fg"], padx=10, pady=6, cursor="hand2",
                  command=pop.destroy).pack(side="left", padx=6)
        pop.columnconfigure(1, weight=1)
        self._popup_smart(pop)

    def _txt_actif(self):
        focused = self.focus_get()
        tab = self._tab()
        for t in tab._corps_pages:
            if t == focused:
                return t
        return tab.corps

    def _copier_page(self):
        txt = self._txt_actif()
        contenu = txt.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(contenu)
        self._statut.set(f"Page {getattr(txt, '_page_num', 1)} copiée dans le presse-papiers.")

    def _copier_tout(self):
        tab = self._tab()
        pages = [t.get("1.0", "end-1c") for t in tab._corps_pages]
        contenu = "\n\n".join(pages)
        self.clipboard_clear()
        self.clipboard_append(contenu)
        n = len(tab._corps_pages)
        self._statut.set(f"{n} page(s) copiées dans le presse-papiers.")

    def _quitter(self):
        for tab_id in self.notebook.tabs():
            tab = self.notebook.nametowidget(tab_id)
            if self._onglet_modifie(tab):
                self.notebook.select(tab)
                rep = messagebox.askyesnocancel(
                    "Quitter",
                    f"Le courrier « {tab.get_titre()} » contient du contenu.\n"
                    "Enregistrer avant de quitter ?")
                if rep is None:
                    return
                if rep:
                    self._sauvegarder_temp(tab)
        self._sauvegarder_expediteur_cache()
        self.destroy()

    def _inserer_date(self):
        self._txt_actif().insert(tk.INSERT, datetime.now().strftime("%d/%m/%Y %H:%M"))

    def _nouveau_courrier(self):
        if messagebox.askyesno("Nouveau", "Effacer le courrier en cours ?"):
            self._tab().vider()
            self.notebook.tab(self._tab(), text="Nouveau courrier")

    # -------------------------------------------------------- modèles

    def _charger_modele(self, nom):
        modele = MODELES[nom]
        tab    = self._tab()
        if tab.corps.get("1.0", "end-1c").strip():
            if not messagebox.askyesno("Modèle",
                                       f"Appliquer le modèle « {nom} » ?\n"
                                       "Le corps actuel sera remplacé."):
                return
        tab.d_objet.set(modele["objet"])
        tab.corps.delete("1.0", "end")
        tab.corps.insert("1.0", modele["corps"])
        self._statut.set(f"Modèle chargé : {nom}")

    # -------------------------------------------------------- historique destinataires

    def _rafraichir_menu_hist(self):
        self._menu_hist.delete(0, "end")
        hist = charger_historique()
        if not hist:
            self._menu_hist.add_command(label="(aucun)", state="disabled")
            return
        for dest in hist:
            label = dest.get("nom", "?")
            if dest.get("societe"):
                label += f" — {dest['societe']}"
            self._menu_hist.add_command(
                label=label,
                command=lambda d=dest: self._tab().remplir_depuis_historique(d))

    # -------------------------------------------------------- format

    def _toggle_wrap(self):
        mode = "word" if self._word_wrap.get() else "none"
        for tab_id in self.notebook.tabs():
            for txt in self.notebook.nametowidget(tab_id)._corps_pages:
                txt.configure(wrap=mode)

    def _toggle_statusbar(self):
        if self._statusbar_v.get():
            self._statusbar.pack(fill="x", side="bottom")
        else:
            self._statusbar.pack_forget()

    def _choisir_police(self):
        PoliceDialog(self)

    def appliquer_police(self):
        f = (self._font_family.get(), self._font_size.get())
        for tab_id in self.notebook.tabs():
            for txt in self.notebook.nametowidget(tab_id)._corps_pages:
                txt.configure(font=f)
        pct = int(self._font_size.get() / 13 * 100)
        self._lbl_zoom.config(text=f"{pct}%")

    def _zoom_plus(self):
        if self._font_size.get() < 36:
            self._font_size.set(self._font_size.get() + 1)
            self.appliquer_police()

    def _zoom_moins(self):
        if self._font_size.get() > 6:
            self._font_size.set(self._font_size.get() - 1)
            self.appliquer_police()

    def _zoom_reset(self):
        self._font_size.set(13)
        self._font_family.set("Courier New")
        self.appliquer_police()

    # -------------------------------------------------------- rechercher / remplacer

    def _rechercher(self):
        if self._search_win and self._search_win.winfo_exists():
            self._search_win.lift(); return
        self._search_win = SearchDialog(self, replace=False)

    def _remplacer(self):
        if self._search_win and self._search_win.winfo_exists():
            self._search_win.lift(); return
        self._search_win = SearchDialog(self, replace=True)

    def _rechercher_suivant(self):
        if self._search_win and self._search_win.winfo_exists():
            self._search_win.suivant()

    def _rechercher_precedent(self):
        if self._search_win and self._search_win.winfo_exists():
            self._search_win.precedent()

    def _aller_a_ligne(self):
        corps = self._tab().corps
        total = int(corps.index("end-1c").split(".")[0])
        pop   = tk.Toplevel(self)
        pop.title("Atteindre")
        pop.resizable(False, False)
        pop.grab_set()
        tk.Label(pop, text=f"Ligne (1 – {total}) :",
                 font=("Courier New", 9), padx=12, pady=8).pack()
        v = tk.IntVar(value=1)
        e = tk.Entry(pop, textvariable=v, font=("Courier New", 11),
                     width=10, justify="center")
        e.pack(padx=12, pady=4)
        e.select_range(0, "end")
        def ok():
            ln = max(1, min(v.get(), total))
            corps.mark_set(tk.INSERT, f"{ln}.0")
            corps.see(f"{ln}.0")
            pop.destroy()
        e.bind("<Return>", lambda _: ok())
        tk.Button(pop, text="OK", font=("Courier New", 11), relief="flat",
                  bg="#2a6099", fg="white", padx=14, pady=4, cursor="hand2",
                  command=ok).pack(pady=(4, 10))
        e.focus_set()
        self._popup_smart(pop)

    def search_in_text(self, terme, case_sensitive=False,
                       start_widget=None, start_pos="1.0"):
        """Cherche dans toutes les pages. Retourne (widget, idx, end) ou (None, None, None)."""
        tab   = self._tab()
        pages = tab._corps_pages
        for t in pages:
            t.tag_remove("trouve", "1.0", "end")
        if not terme:
            return None, None, None
        count = tk.IntVar()
        if start_widget is None or start_widget not in pages:
            start_widget = pages[0]
            start_pos    = "1.0"
        si = pages.index(start_widget)
        # 1. Page de départ à partir de start_pos
        idx = start_widget.search(terme, start_pos, stopindex="end",
                                  count=count, nocase=not case_sensitive)
        if idx:
            end = f"{idx}+{count.get()}c"
            start_widget.tag_add("trouve", idx, end)
            start_widget.see(idx)
            return start_widget, idx, end
        # 2. Pages suivantes
        for t in pages[si + 1:]:
            idx = t.search(terme, "1.0", stopindex="end",
                           count=count, nocase=not case_sensitive)
            if idx:
                end = f"{idx}+{count.get()}c"
                t.tag_add("trouve", idx, end)
                t.see(idx)
                return t, idx, end
        # 3. Wraparound depuis le début jusqu'à start_pos
        for i, t in enumerate(pages[:si + 1]):
            stop = start_pos if i == si else "end"
            idx = t.search(terme, "1.0", stopindex=stop,
                           count=count, nocase=not case_sensitive)
            if idx:
                end = f"{idx}+{count.get()}c"
                t.tag_add("trouve", idx, end)
                t.see(idx)
                return t, idx, end
        return None, None, None

    # -------------------------------------------------------- LanguageTool

    def _detect_lt(self):
        if not HAS_REQUESTS:
            self._statut.set("Module 'requests' manquant — correcteur désactivé")
            return
        try:
            r = requests.get(LT_LOCAL.replace("/check", "/languages"), timeout=1)
            if r.status_code == 200:
                self._lt_url = LT_LOCAL
                self._statut.set("LanguageTool local connecté")
                return
        except Exception:
            pass
        self._lt_url = LT_PUBLIC
        self._statut.set("LanguageTool : API publique")

    def _verifier(self):
        if not self._lt_url:
            messagebox.showwarning("Correcteur", "Correcteur non disponible.")
            return
        tab = self._tab()
        tab._effacer_erreurs()
        pages_a_verifier = [(txt, txt.get("1.0", "end-1c"))
                            for txt in tab._corps_pages
                            if txt.get("1.0", "end-1c").strip()]
        if not pages_a_verifier:
            return
        self._statut.set("Vérification en cours…")
        self.update_idletasks()
        for txt, texte in pages_a_verifier:
            threading.Thread(target=self._appel_lt,
                             args=(tab, txt, texte), daemon=True).start()

    def _appel_lt(self, tab, txt, texte):
        try:
            r = requests.post(self._lt_url,  # type: ignore[union-attr]
                              data={"text": texte, "language": "fr"},
                              timeout=15)
            matches = r.json().get("matches", [])
            self.after(0, self._afficher_erreurs, tab, txt, texte, matches)
        except Exception as e:
            self.after(0, self._statut.set, f"Erreur LT : {e}")

    def _afficher_erreurs(self, tab, txt, texte, matches):
        tab.afficher_erreurs(txt, texte, matches)
        nb = sum(len(v) for v in tab._match_ranges.values())
        if nb:
            self._statut.set(
                f"{nb} erreur{'s' if nb > 1 else ''}"
                " — clic droit sur un mot souligné pour corriger")
        else:
            self._statut.set("Aucune erreur détectée.")


# ================================================================ DIALOGUE RECHERCHER

class SearchDialog(tk.Toplevel):
    def __init__(self, app, replace=False):
        super().__init__(app)
        self.app     = app
        self.replace = replace
        self.title("Remplacer" if replace else "Rechercher")
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self._last_widget = None
        self._last_pos    = "1.0"
        self._build()
        app._popup_smart(self)

    def _build(self):
        pad = dict(padx=10, pady=4)
        tk.Label(self, text="Rechercher :", font=("Courier New", 9), anchor="w").grid(
            row=0, column=0, sticky="w", **pad)
        self.v_terme = tk.StringVar()
        tk.Entry(self, textvariable=self.v_terme, font=("Courier New", 11),
                 width=30).grid(row=0, column=1, columnspan=2, sticky="ew", **pad)
        if self.replace:
            tk.Label(self, text="Remplacer par :", font=("Courier New", 9), anchor="w").grid(
                row=1, column=0, sticky="w", **pad)
            self.v_remplacement = tk.StringVar()
            tk.Entry(self, textvariable=self.v_remplacement, font=("Courier New", 11),
                     width=30).grid(row=1, column=1, columnspan=2, sticky="ew", **pad)
        self.v_case = tk.BooleanVar(value=False)
        tk.Checkbutton(self, text="Respecter la casse", variable=self.v_case,
                       font=("Courier New", 9)).grid(row=2, column=0, columnspan=2,
                                                     sticky="w", padx=10)
        r = 3
        tk.Button(self, text="Suivant", font=("Courier New", 11), width=12,
                  relief="flat", bg="#2a6099", fg="white", cursor="hand2",
                  command=self.suivant).grid(row=r, column=0, **pad)
        tk.Button(self, text="Précédent", font=("Courier New", 11), width=12,
                  relief="flat", bg="#eceae4", cursor="hand2",
                  command=self.precedent).grid(row=r, column=1, **pad)
        if self.replace:
            tk.Button(self, text="Remplacer", font=("Courier New", 11), width=12,
                      relief="flat", bg="#555", fg="white", cursor="hand2",
                      command=self._remplacer_un).grid(row=r+1, column=0, **pad)
            tk.Button(self, text="Remplacer tout", font=("Courier New", 11), width=12,
                      relief="flat", bg="#333", fg="white", cursor="hand2",
                      command=self._remplacer_tout).grid(row=r+1, column=1, **pad)
        tk.Button(self, text="Fermer", font=("Courier New", 9), relief="flat",
                  bg="#eceae4", cursor="hand2",
                  command=self.destroy).grid(row=r+2, column=1, sticky="e", **pad)

    def suivant(self):
        terme = self.v_terme.get()
        if not terme: return
        w, idx, end = self.app.search_in_text(
            terme, self.v_case.get(),
            start_widget=self._last_widget, start_pos=self._last_pos)
        if idx:
            self._last_widget = w
            self._last_pos    = end
        else:
            self._last_widget = None
            self._last_pos    = "1.0"
            messagebox.showinfo("Rechercher", f"« {terme} » introuvable.", parent=self)

    def precedent(self):
        terme = self.v_terme.get()
        if not terme: return
        tab   = self.app._tab()
        pages = tab._corps_pages
        for t in pages:
            t.tag_remove("trouve", "1.0", "end")
        count = tk.IntVar()
        if self._last_widget is None or self._last_widget not in pages:
            self._last_widget = pages[-1]
            self._last_pos    = "end"
        si = pages.index(self._last_widget)

        def _chercher_backward(widget, stop):
            idx = widget.search(terme, "1.0", stopindex=stop,
                                nocase=not self.v_case.get(), backwards=True)
            if idx:
                widget.search(terme, idx, count=count)
                end = f"{idx}+{count.get()}c"
                widget.tag_add("trouve", idx, end)
                widget.see(idx)
                self._last_widget = widget
                self._last_pos    = idx
                return True
            return False

        if _chercher_backward(self._last_widget, self._last_pos): return
        for t in reversed(pages[:si]):
            if _chercher_backward(t, "end"): return
        for t in reversed(pages[si:]):
            if _chercher_backward(t, "end"): return
        messagebox.showinfo("Rechercher", f"« {terme} » introuvable.", parent=self)

    def _remplacer_un(self):
        tab = self.app._tab()
        for t in tab._corps_pages:
            ranges = t.tag_ranges("trouve")
            if ranges:
                t.delete(ranges[0], ranges[1])
                t.insert(ranges[0], self.v_remplacement.get())
                self._last_widget = t
                self._last_pos    = str(ranges[0])
                self.suivant()
                return
        self.suivant()

    def _remplacer_tout(self):
        terme = self.v_terme.get()
        rempl = self.v_remplacement.get()
        if not terme: return
        tab   = self.app._tab()
        flags = 0 if self.v_case.get() else re.IGNORECASE
        total = 0
        for t in tab._corps_pages:
            contenu = t.get("1.0", "end-1c")
            nouveau, nb = re.subn(re.escape(terme), rempl, contenu, flags=flags)
            if nb:
                t.delete("1.0", "end")
                t.insert("1.0", nouveau)
                total += nb
        if total:
            messagebox.showinfo("Remplacer", f"{total} remplacement(s).", parent=self)
        else:
            messagebox.showinfo("Remplacer", f"« {terme} » introuvable.", parent=self)


# ================================================================ DIALOGUE POLICE

class PoliceDialog(tk.Toplevel):
    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Police")
        self.resizable(False, False)
        self.grab_set()
        self._build()
        app._popup_smart(self)

    def _build(self):
        pad = dict(padx=10, pady=4)
        tk.Label(self, text="Police :", font=("Courier New", 9)).grid(
            row=0, column=0, sticky="w", **pad)
        self.v_famille = tk.StringVar(value=self.app._font_family.get())
        familles = sorted(set(tkfont.families()))
        lf = tk.Frame(self); lf.grid(row=1, column=0, sticky="nsew", **pad)
        sb = tk.Scrollbar(lf); sb.pack(side="right", fill="y")
        self.lb = tk.Listbox(lf, font=("Courier New", 9), width=32, height=12,
                             yscrollcommand=sb.set, exportselection=False)
        sb.config(command=self.lb.yview)
        self.lb.pack(side="left")
        for f in familles:
            self.lb.insert("end", f)
        if self.app._font_family.get() in familles:
            idx = familles.index(self.app._font_family.get())
            self.lb.selection_set(idx); self.lb.see(idx)
        self.lb.bind("<<ListboxSelect>>", self._apercu)
        tk.Label(self, text="Taille :", font=("Courier New", 9)).grid(
            row=0, column=1, sticky="w", **pad)
        self.v_taille = tk.IntVar(value=self.app._font_size.get())
        tk.Spinbox(self, from_=6, to=36, textvariable=self.v_taille, font=("Courier New", 11),
                   width=5, command=self._apercu).grid(row=1, column=1, sticky="nw", **pad)
        tk.Label(self, text="Aperçu :", font=("Courier New", 9)).grid(
            row=2, column=0, sticky="w", **pad)
        self._lbl_ap = tk.Label(self, text="Le vif renard brun…",
                                font=(self.app._font_family.get(), self.app._font_size.get()),
                                bg="white", relief="solid", width=36, pady=8)
        self._lbl_ap.grid(row=3, column=0, columnspan=2, **pad)
        bf = tk.Frame(self); bf.grid(row=4, column=0, columnspan=2, sticky="e", **pad)
        tk.Button(bf, text="OK", font=("Courier New", 11), relief="flat",
                  bg="#2a6099", fg="white", padx=14, pady=4, cursor="hand2",
                  command=self._ok).pack(side="right", padx=4)
        tk.Button(bf, text="Annuler", font=("Courier New", 9), relief="flat",
                  bg="#eceae4", padx=10, pady=4, cursor="hand2",
                  command=self.destroy).pack(side="right")

    def _apercu(self, _=None):
        sel = self.lb.curselection()
        if sel: self.v_famille.set(self.lb.get(sel[0]))
        try:
            self._lbl_ap.config(font=(self.v_famille.get(), self.v_taille.get()))
        except Exception:
            pass

    def _ok(self):
        self.app._font_family.set(self.v_famille.get())
        self.app._font_size.set(self.v_taille.get())
        self.app.appliquer_police()
        self.destroy()


def _associer_extension():
    """Associe .dactime a DacTime dans HKCU (sans droits admin). Idempotent."""
    import sys, winreg
    exe = sys.executable
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\.dactime") as k:
            winreg.SetValue(k, "", winreg.REG_SZ, "DacTime.Document")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\DacTime.Document") as k:
            winreg.SetValue(k, "", winreg.REG_SZ, "Courrier DacTime")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                              r"Software\Classes\DacTime.Document\DefaultIcon") as k:
            winreg.SetValue(k, "", winreg.REG_SZ, f'"{exe}",0')
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                              r"Software\Classes\DacTime.Document\shell\open\command") as k:
            winreg.SetValue(k, "", winreg.REG_SZ, f'"{exe}" "%1"')
    except Exception:
        pass


if __name__ == "__main__":
    import sys as _sys
    _associer_extension()
    app = CourrierApp()
    if len(_sys.argv) > 1 and os.path.isfile(_sys.argv[1]):
        app.after(100, lambda: app._charger_fichier(_sys.argv[1]))
    app.mainloop()
