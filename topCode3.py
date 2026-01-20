import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
from datetime import datetime

DB_NAME = 'parts_database.db'

def create_database():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS parts (
            part_number TEXT NOT NULL,
            revision   TEXT NOT NULL,
            description TEXT,
            where_used  TEXT,
            status      TEXT DEFAULT 'Active',
            folder_path TEXT,
            file_names  TEXT,
            last_updated TEXT,
            PRIMARY KEY (part_number, revision)
        )
    ''')
    conn.commit()
    conn.close()

def insert_sample_data():
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    sample_data = [
        ('PN001', 'A', 'Widget Assembly Base Unit', 'Assy-101, Assy-202', 'Active', r'C:\Projects\WidgetLine', 'base.dwg, BOM.xlsx, spec.pdf', now),
        ('PN002', 'B', 'Control Module v2',        'Product X, Product Y Rev C', 'Active', r'C:\Designs\Controls', 'module_v2.step, wiring.pdf', now),
        ('PN003', 'A', 'Sensor Housing Stainless', 'Assy-303, Test Fixture 4', 'Active', r'C:\Parts\ Housings', 'housing_v1.igs, 003_revA.pdf', now),
        ('PN004', 'C', 'Gearbox 90deg 1:5 ratio',  'Robot Arm v3, Conveyor B', 'Obsolete', '', 'old_gearbox.stp', now),
        ('PN005', 'A', 'Power Supply 24V 5A',      'Control Cabinet A, Backup Unit', 'Active', r'C:\Purchased\PSU', 'MEANWELL_LRS-100.pdf', now),
        ('PN006', 'A', 'Mounting Bracket Left',    'Assy-101, Assy-202', 'New', r'C:\Brackets', 'left_bracket.SLDPRT', now),
        ('PN007', 'B', 'Mounting Bracket Right',   'Assy-101, Assy-202', 'Active', r'C:\Brackets', 'right_bracket_revB.SLDPRT, DXF_export.dxf', now),
        ('PN101', '1', 'Main Chassis Weldment',    'Product X Final, Product Y', 'Active', r'C:\Weldments\Main', 'chassis_v1.dwg', now),
    ]
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.executemany('''
        INSERT OR REPLACE INTO parts
        (part_number, revision, description, where_used, status, folder_path, file_names, last_updated)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', sample_data)
    conn.commit()
    conn.close()

class PartsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Parts Database Manager")
        self.root.geometry("1400x750")
        self.sort_column = 'part_number'
        self.sort_reverse = False

        # ── Search / Filter Frame ───────────────────────────────────────
        filter_frame = ttk.LabelFrame(root, text="Search & Filter", padding=10)
        filter_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=5, columnspan=2)

        self.search_entries = {}
        fields = ['part_number', 'revision', 'description', 'where_used', 'status', 'folder_path', 'file_names']
        for i, field in enumerate(fields):
            row = i // 2
            col = (i % 2) * 4
            ttk.Label(filter_frame, text=field.replace('_', ' ').title() + ":").grid(row=row, column=col, padx=5, pady=3, sticky="e")
            entry = ttk.Entry(filter_frame, width=24)
            entry.grid(row=row, column=col+1, padx=5, pady=3, sticky="w")
            self.search_entries[field] = entry

        btn_frame = ttk.Frame(filter_frame)
        btn_frame.grid(row=len(fields)//2 + 1, column=0, columnspan=8, pady=10, sticky="w")
        ttk.Button(btn_frame, text="Apply Filter", command=self.apply_filter).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Clear Filter", command=self.clear_filter).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Show All",     command=self.show_all).pack(side="left", padx=5)

        # ── Action Buttons ──────────────────────────────────────────────
        action_frame = ttk.Frame(root)
        action_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5, columnspan=2)
        ttk.Button(action_frame, text="Add New Part",   command=self.add_record).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Edit Selected",  command=self.edit_record).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Delete Selected",command=self.delete_records).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Export to CSV",  command=self.export_to_csv).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Import from CSV",command=self.import_from_csv).pack(side="left", padx=5)

        # ── Treeview ────────────────────────────────────────────────────
        columns = ('Part Number', 'Rev', 'Description', 'Where Used', 'Status', 'Folder Path', 'File Names', 'Last Updated')
        self.tree = ttk.Treeview(root, columns=columns, show='headings')

        col_info = {
            'Part Number':  ('part_number', 110),
            'Rev':          ('revision',     60),
            'Description':  ('description',  220),
            'Where Used':   ('where_used',   160),
            'Status':       ('status',       90),
            'Folder Path':  ('folder_path',  180),
            'File Names':   ('file_names',   220),
            'Last Updated': ('last_updated', 140),
        }

        for col in columns:
            db_col, width = col_info[col]
            self.tree.heading(col, text=col, command=lambda c=db_col: self.sort_by_column(c))
            self.tree.column(col, width=width, stretch=True, anchor=tk.W)

        self.tree.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        scrollbar = ttk.Scrollbar(root, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=2, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind("<Double-1>", self.show_details)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)

        # Color tags
        self.tree.tag_configure('New',     foreground='dark green')
        self.tree.tag_configure('Active',  foreground='navy')
        self.tree.tag_configure('Obsolete',foreground='maroon')

        self.current_where = ""
        self.current_params = ()
        self.show_all()

    def sort_by_column(self, col):
        if self.sort_column == col:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_reverse = False
            self.sort_column = col
        self.refresh_view()

    def refresh_view(self):
        self.load_data(self.current_where, self.current_params)

    def load_data(self, where_clause="", params=()):
        self.current_where = where_clause
        self.current_params = params

        for item in self.tree.get_children():
            self.tree.delete(item)

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        order_by = f"ORDER BY {self.sort_column} COLLATE NOCASE {'DESC' if self.sort_reverse else 'ASC'}"
        query = f"""
            SELECT part_number, revision, description, where_used, status,
                   folder_path, file_names, last_updated
            FROM parts {where_clause} {order_by}
        """
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()

        # Update sort arrows
        col_info = {
            'Part Number': 'part_number', 'Rev': 'revision', 'Description': 'description',
            'Where Used': 'where_used', 'Status': 'status', 'Folder Path': 'folder_path',
            'File Names': 'file_names', 'Last Updated': 'last_updated'
        }
        for display_col in self.tree['columns']:
            db_col = col_info.get(display_col, '')
            arrow = ""
            if db_col == self.sort_column:
                arrow = " ↓" if self.sort_reverse else " ↑"
            self.tree.heading(display_col, text=display_col + arrow)

        for row in rows:
            tag = row[4] if row[4] in ['New', 'Active', 'Obsolete'] else ''
            self.tree.insert('', tk.END, values=row, tags=(tag,))

    def build_filter(self):
        conditions = []
        params = []
        fields = ['part_number', 'revision', 'description', 'where_used', 'status', 'folder_path', 'file_names']
        for field in fields:
            term = self.search_entries[field].get().strip()
            if term:
                conditions.append(f"{field} LIKE ?")
                params.append(f"%{term}%")
        if conditions:
            return "WHERE " + " AND ".join(conditions), tuple(params)
        return "", ()

    def apply_filter(self):
        where, params = self.build_filter()
        self.load_data(where, params)

    def clear_filter(self):
        for entry in self.search_entries.values():
            entry.delete(0, tk.END)

    def show_all(self):
        self.clear_filter()
        self.load_data()

    def get_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a row first.")
            return None
        return self.tree.item(selected[0])['values']

    def add_record(self):
        self.open_edit_window(is_new=True)

    def edit_record(self):
        values = self.get_selected()
        if values:
            self.open_edit_window(is_new=False, values=values)

    def open_edit_window(self, is_new=True, values=None):
        win = tk.Toplevel(self.root)
        win.title("New Part" if is_new else "Edit Part")
        win.geometry("700x600")
        win.transient(self.root)
        win.grab_set()

        frame = ttk.Frame(win, padding=15)
        frame.pack(fill="both", expand=True)

        entries = {}
        fields = ['part_number', 'revision', 'description', 'where_used', 'status', 'folder_path', 'file_names']
        labels = ['Part Number*', 'Revision*', 'Description', 'Where Used', 'Status', 'Folder Path', 'File Names (comma sep.)']

        for i, (field, label) in enumerate(zip(fields, labels)):
            ttk.Label(frame, text=label).grid(row=i, column=0, sticky="e", pady=6, padx=8)
            if field == 'status':
                combo = ttk.Combobox(frame, values=['New', 'Active', 'Obsolete'], width=35)
                combo.grid(row=i, column=1, sticky="w", pady=6)
                entries[field] = combo
            else:
                entry = ttk.Entry(frame, width=60)
                entry.grid(row=i, column=1, sticky="ew", pady=6)
                entries[field] = entry

        if not is_new and values:
            for field, val in zip(fields, values[:7]):
                if field == 'status':
                    entries[field].set(val)
                else:
                    entries[field].insert(0, val or "")
            entries['part_number'].config(state='disabled')
            entries['revision'].config(state='disabled')

        def save():
            data = [entries[f].get().strip() for f in fields]
            if not data[0] or not data[1]:
                messagebox.showerror("Error", "Part Number and Revision are required.")
                return

            now = datetime.now().strftime('%Y-%m-%d %H:%M')
            data.append(now)  # last_updated

            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            try:
                if is_new:
                    cursor.execute('''
                        INSERT INTO parts
                        (part_number, revision, description, where_used, status, folder_path, file_names, last_updated)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', data)
                else:
                    cursor.execute('''
                        UPDATE parts SET
                            description = ?, where_used = ?, status = ?,
                            folder_path = ?, file_names = ?, last_updated = ?
                        WHERE part_number = ? AND revision = ?
                    ''', (data[2], data[3], data[4], data[5], data[6], data[7], data[0], data[1]))

                conn.commit()
                win.destroy()
                self.refresh_view()
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Part Number + Revision combination already exists.")
            finally:
                conn.close()

        ttk.Button(frame, text="Save", command=save).grid(row=len(fields)+1, column=0, columnspan=2, pady=20)
        ttk.Button(frame, text="Cancel", command=win.destroy).grid(row=len(fields)+2, column=0, columnspan=2)

    def delete_records(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Select at least one row to delete.")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete {len(selected)} record(s)?"):
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        for item in selected:
            values = self.tree.item(item)['values']
            cursor.execute("DELETE FROM parts WHERE part_number = ? AND revision = ?", (values[0], values[1]))
        conn.commit()
        conn.close()
        self.refresh_view()

    def show_details(self, event):
        values = self.get_selected()
        if not values:
            return

        win = tk.Toplevel(self.root)
        win.title(f"Details: {values[0]} Rev {values[1]}")
        win.geometry("800x600")

        frame = ttk.Frame(win, padding=20)
        frame.pack(fill="both", expand=True)

        labels = ['Part Number', 'Revision', 'Description', 'Where Used', 'Status', 'Folder Path', 'File Names', 'Last Updated']
        for i, (lbl, val) in enumerate(zip(labels, values)):
            ttk.Label(frame, text=lbl + ":", font=("Arial", 10, "bold")).grid(row=i, column=0, sticky="ne", padx=10, pady=6)
            ttk.Label(frame, text=val or "—", wraplength=600, justify="left").grid(row=i, column=1, sticky="nw", pady=6)

        ttk.Button(frame, text="Close", command=win.destroy).grid(row=len(labels)+1, column=0, columnspan=2, pady=20)

    def export_to_csv(self):
        if not self.tree.get_children():
            messagebox.showinfo("Nothing to export", "No records to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        with open(file_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Part Number', 'Revision', 'Description', 'Where Used', 'Status', 'Folder Path', 'File Names', 'Last Updated'])
            for item in self.tree.get_children():
                writer.writerow(self.tree.item(item)['values'])

        messagebox.showinfo("Export Complete", f"Saved to:\n{file_path}")

    def import_from_csv(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Select CSV file to import"
        )
        if not file_path:
            return

        inserted = updated = skipped = errors = 0
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                header_skipped = False
                for row in reader:
                    if not row or len(row) < 2:
                        skipped += 1
                        continue
                    if not header_skipped and 'part' in row[0].lower() and 'revision' in row[1].lower():
                        header_skipped = True
                        continue

                    res = self._process_csv_row(cursor, row)
                    inserted += res['inserted']
                    updated   += res['updated']
                    skipped   += res['skipped']
                    errors    += res['error']

            conn.commit()
            msg = "Import finished.\n\n"
            if inserted: msg += f"Inserted: {inserted}\n"
            if updated:  msg += f"Updated: {updated}\n"
            if skipped:  msg += f"Skipped: {skipped}\n"
            if errors:   msg += f"Errors: {errors}\n"

            if inserted + updated > 0:
                self.refresh_view()
                messagebox.showinfo("Import Summary", msg)
            else:
                messagebox.showwarning("Import Summary", msg + "\nNo records were imported.")
        except Exception as e:
            conn.rollback()
            messagebox.showerror("Import Error", f"Failed to import file:\n{str(e)}")
        finally:
            conn.close()

    def _process_csv_row(self, cursor, row):
        result = {'inserted':0, 'updated':0, 'skipped':0, 'error':0}
        try:
            # Pad row if too short
            while len(row) < 8:
                row.append('')

            part_number = row[0].strip()
            revision    = row[1].strip()
            if not part_number or not revision:
                result['skipped'] = 1
                return result

            description = row[2].strip() or None
            where_used  = row[3].strip() or None
            status      = row[4].strip() or 'Active'
            folder_path = row[5].strip() or None
            file_names  = row[6].strip() or None
            last_updated = row[7].strip() or datetime.now().strftime('%Y-%m-%d %H:%M')

            if status not in ['New', 'Active', 'Obsolete']:
                status = 'Active'

            cursor.execute('''
                INSERT OR REPLACE INTO parts
                (part_number, revision, description, where_used, status, folder_path, file_names, last_updated)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (part_number, revision, description, where_used, status, folder_path, file_names, last_updated))

            if cursor.rowcount == 1:
                result['inserted'] = 1
            else:
                result['updated'] = 1
        except Exception:
            result['error'] = 1
        return result

if __name__ == "__main__":
    create_database()
    insert_sample_data()
    root = tk.Tk()
    app = PartsApp(root)
    root.mainloop()