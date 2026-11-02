"""
gestione_sostituzioni.editor
EditorWindow: finestra di modifica CSV integrata con protezione anti-troncamento.
Grafica premium con zebra-striping, intestazioni colorate e toolbar moderna.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import csv
import io
import sys
import logging

logger = logging.getLogger("GestioneSostituzioni.Editor")

# ── Palette colori editor ──────────────────────────────────────────────
_EDITOR_COLORS = {
    'toolbar_bg':     '#1E293B',   # Slate scuro per la toolbar
    'toolbar_fg':     '#F1F5F9',   # Testo chiaro sulla toolbar
    'toolbar_accent': '#38BDF8',   # Azzurro cielo per accenti
    'btn_save_bg':    '#22C55E',   # Verde vivace per Salva
    'btn_save_fg':    '#FFFFFF',
    'btn_save_hover': '#16A34A',
    'heading_bg':     '#334155',   # Slate medio per le intestazioni colonne
    'heading_fg':     '#F8FAFC',
    'row_even':       '#FFFFFF',   # Bianco per righe pari
    'row_odd':        '#F1F5F9',   # Grigio ghiaccio per righe dispari
    'row_selected':   '#DBEAFE',   # Azzurro tenue per selezione
    'row_sel_fg':     '#1E3A5F',
    'cell_border':    '#CBD5E1',   # Bordo sottile tra celle
    'status_bg':      '#F8FAFC',   # Sfondo barra di stato
    'status_fg':      '#64748B',   # Testo grigio nella status bar
    'overflow_bg':    '#FEF3C7',   # Giallo caldo per avviso overflow
    'overflow_fg':    '#92400E',
    'edit_bg':        '#FFFBEB',   # Sfondo cella in editing
    'edit_border':    '#F59E0B',   # Bordo cella in editing
    'window_bg':      '#F1F5F9',   # Sfondo finestra
}


class EditorWindow(tk.Toplevel):
    """Editor CSV ottimizzato con grafica premium e protezione overflow righe."""

    def __init__(self, parent, title, content, on_save, max_rows=1000):
        super().__init__(parent)
        self.title(f"📊 Editor — {title}")
        self.geometry("1400x800")
        self.configure(bg=_EDITOR_COLORS['window_bg'])
        self.on_save = on_save
        self.max_rows = max_rows
        self.overflow_rows = []

        # ── Stile Treeview (una sola volta, leggero) ───────────────────
        self._apply_style()

        # ── Toolbar moderna ────────────────────────────────────────────
        toolbar = tk.Frame(self, bg=_EDITOR_COLORS['toolbar_bg'], padx=12, pady=8)
        toolbar.pack(fill="x")

        # Pulsante Salva
        self.btn_save = tk.Button(
            toolbar, text="💾  SALVA E CHIUDI",
            bg=_EDITOR_COLORS['btn_save_bg'], fg=_EDITOR_COLORS['btn_save_fg'],
            activebackground=_EDITOR_COLORS['btn_save_hover'],
            font=("Helvetica", 11, "bold"), relief="flat",
            padx=16, pady=4, cursor="hand2", command=self.save,
            highlightbackground=_EDITOR_COLORS['btn_save_bg']
        )
        self.btn_save.pack(side="left")

        # Hint editing
        hint_frame = tk.Frame(toolbar, bg=_EDITOR_COLORS['toolbar_bg'])
        hint_frame.pack(side="left", padx=20)
        tk.Label(hint_frame, text="✏️", font=("Segoe UI", 12),
                 bg=_EDITOR_COLORS['toolbar_bg']).pack(side="left")
        tk.Label(hint_frame,
                 text="Doppio click su una cella per modificare  •  Invio per confermare  •  Esc per annullare",
                 bg=_EDITOR_COLORS['toolbar_bg'], fg=_EDITOR_COLORS['toolbar_accent'],
                 font=("Segoe UI", 10)).pack(side="left", padx=6)

        # ── Avviso overflow (se presente, verrà mostrato dopo il parsing) ──
        self._overflow_frame = tk.Frame(toolbar, bg=_EDITOR_COLORS['toolbar_bg'])
        self._overflow_frame.pack(side="right")

        # ── Parse content ──────────────────────────────────────────────
        self.delimiter = ';'
        self.rows = []
        if content:
            first_line = content.splitlines()[0] if content.splitlines() else ""
            if first_line.count(';') >= first_line.count(','):
                self.delimiter = ';'
            else:
                self.delimiter = ','
            try:
                reader = csv.reader(content.splitlines(), delimiter=self.delimiter)
                self.rows = list(reader)
            except Exception as e:
                logger.error(f"Errore parsing contenuto editor: {e}")
                self.rows = []

        if not self.rows:
            self.rows = [["" for _ in range(60)] for _ in range(100)]

        # ── Treeview setup ─────────────────────────────────────────────
        max_cols = 0
        for r in self.rows:
            max_cols = max(max_cols, len(r))
        max_cols = max(max_cols, 20)
        max_cols = min(max_cols, 100)

        # Contenitore con bordo sottile
        tree_container = tk.Frame(self, bg=_EDITOR_COLORS['cell_border'], bd=1, relief="solid")
        tree_container.pack(fill="both", expand=True, padx=8, pady=(4, 0))

        self.tree = ttk.Treeview(
            tree_container,
            columns=[str(i) for i in range(max_cols)],
            show="headings",
            style="Editor.Treeview"
        )

        # Scrollbars
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        # Intestazioni e colonne
        for i in range(max_cols):
            self.tree.heading(str(i), text=f"  Col {i+1}  ")
            self.tree.column(str(i), width=180, minwidth=80, stretch=False)

        # ── Zebra-striping con tag (applicati durante l'insert, costo zero) ──
        self.tree.tag_configure("row_even", background=_EDITOR_COLORS['row_even'])
        self.tree.tag_configure("row_odd", background=_EDITOR_COLORS['row_odd'])

        for i, row in enumerate(self.rows[:self.max_rows]):
            vals = row + [""] * (max_cols - len(row))
            tag = "row_even" if i % 2 == 0 else "row_odd"
            self.tree.insert("", "end", iid=str(i), values=vals, tags=(tag,))

        # ── Overflow warning ───────────────────────────────────────────
        if len(self.rows) > self.max_rows:
            self.overflow_rows = self.rows[self.max_rows:]
            overflow_pill = tk.Frame(self._overflow_frame, bg=_EDITOR_COLORS['overflow_bg'],
                                     padx=10, pady=2, highlightbackground=_EDITOR_COLORS['overflow_fg'],
                                     highlightthickness=1)
            overflow_pill.pack(side="right")
            tk.Label(overflow_pill,
                     text=f"⚠ {len(self.overflow_rows)} righe nascoste (preservate al salvataggio)",
                     bg=_EDITOR_COLORS['overflow_bg'], fg=_EDITOR_COLORS['overflow_fg'],
                     font=("Segoe UI", 9, "bold")).pack()
            logger.info(f"Editor: {len(self.overflow_rows)} righe in overflow (preservate)")

        # ── Barra di stato ─────────────────────────────────────────────
        status_bar = tk.Frame(self, bg=_EDITOR_COLORS['status_bg'], padx=12, pady=4)
        status_bar.pack(fill="x", side="bottom")

        rows_shown = min(len(self.rows), self.max_rows)
        total_rows = len(self.rows)
        self.lbl_status = tk.Label(
            status_bar,
            text=f"📄 {rows_shown} righe visualizzate  •  {max_cols} colonne  •  {total_rows} righe totali  •  Delimitatore: «{self.delimiter}»",
            bg=_EDITOR_COLORS['status_bg'], fg=_EDITOR_COLORS['status_fg'],
            font=("Segoe UI", 9)
        )
        self.lbl_status.pack(side="left")

        self.lbl_file_title = tk.Label(
            status_bar, text=title,
            bg=_EDITOR_COLORS['status_bg'], fg=_EDITOR_COLORS['heading_bg'],
            font=("Segoe UI", 9, "bold")
        )
        self.lbl_file_title.pack(side="right")

        # ── Bind eventi ────────────────────────────────────────────────
        self.tree.bind("<Double-1>", self.on_double_click)

    def _apply_style(self):
        """Applica lo stile personalizzato al Treeview dell'editor."""
        style = ttk.Style(self)

        style.configure("Editor.Treeview",
                         background=_EDITOR_COLORS['row_even'],
                         foreground="#1E293B",
                         rowheight=28,
                         font=("Segoe UI", 10),
                         borderwidth=0,
                         fieldbackground=_EDITOR_COLORS['row_even'])

        style.configure("Editor.Treeview.Heading",
                         background=_EDITOR_COLORS['heading_bg'],
                         foreground=_EDITOR_COLORS['heading_fg'],
                         font=("Segoe UI", 10, "bold"),
                         relief="flat",
                         padding=(6, 4))

        style.map("Editor.Treeview",
                   background=[("selected", _EDITOR_COLORS['row_selected'])],
                   foreground=[("selected", _EDITOR_COLORS['row_sel_fg'])])

        style.map("Editor.Treeview.Heading",
                   background=[("active", "#475569")])

        # Layout senza bordi interni
        style.layout("Editor.Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])

    def on_double_click(self, event):
        """Gestisce il doppio click su una cella per l'editing inline."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        col_id = self.tree.identify_column(event.x)
        row_id = self.tree.identify_row(event.y)
        if not col_id or not row_id:
            return

        col_idx = int(col_id.replace("#", "")) - 1

        # Get bounds
        try:
            x, y, w, h = self.tree.bbox(row_id, col_id)
        except (ValueError, tk.TclError):
            return

        # Current val
        vals = list(self.tree.item(row_id, "values"))
        curr_val = vals[col_idx] if col_idx < len(vals) else ""

        # Create styled entry
        entry = tk.Entry(self.tree, width=20,
                         font=("Segoe UI", 10),
                         bg=_EDITOR_COLORS['edit_bg'],
                         fg="#1E293B",
                         insertbackground="#1E293B",
                         highlightcolor=_EDITOR_COLORS['edit_border'],
                         highlightbackground=_EDITOR_COLORS['edit_border'],
                         highlightthickness=2,
                         relief="flat",
                         bd=0)
        entry.place(x=x, y=y, width=max(w, 120), height=h)
        entry.insert(0, curr_val)
        entry.select_range(0, "end")
        entry.focus_set()

        def commit(e=None):
            new_val = entry.get()
            vals = list(self.tree.item(row_id, "values"))
            while len(vals) <= col_idx:
                vals.append("")
            vals[col_idx] = new_val
            self.tree.item(row_id, values=vals)
            entry.destroy()

        def cancel(e=None):
            entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)

    def save(self):
        """Salva il contenuto dell'editor, incluse le righe in overflow."""
        output = io.StringIO()
        try:
            writer = csv.writer(output, delimiter=self.delimiter, quoting=csv.QUOTE_MINIMAL)

            for child in self.tree.get_children():
                vals = self.tree.item(child, "values")
                row = list(vals)
                while row and not row[-1]:
                    row.pop()
                if any(row):
                    writer.writerow([str(x) for x in row])

            # Riaccodare le righe in overflow che non erano visualizzate
            for row in self.overflow_rows:
                row_copy = list(row)
                while row_copy and not row_copy[-1]:
                    row_copy.pop()
                if any(row_copy):
                    writer.writerow([str(x) for x in row_copy])

            new_content = output.getvalue()
            self.on_save(new_content)
            self.destroy()
            logger.info("Editor: contenuto salvato con successo")
        except Exception as e:
            logger.error(f"Errore salvataggio editor: {e}", exc_info=True)
            messagebox.showerror("Errore Salvataggio", f"Impossibile salvare il contenuto CSV:\n{e}")
