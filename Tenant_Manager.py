import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.simpledialog as simpledialog
import tkinter.font as tkfont
import pandas as pd
from datetime import datetime
import os
import tempfile
import webbrowser
import uuid


def current_nepali_year():
    """
    Auto Nepali year based on Baisakh (mid-April).
    Rule:
    - Before April 14 → AD + 56
    - On/After April 14 → AD + 57
    """
    today = datetime.now()
    ad_year = today.year

    # Baisakh approx: April 14
    if (today.month, today.day) >= (4, 14):
        return ad_year + 57
    else:
        return ad_year + 56


def current_nepali_date_str():
    """
    Auto Nepali date (YYYY-MM-DD approx).
    Day/month are placeholders but year is correct.
    """
    return f"{current_nepali_year()}-01-01"


# Nepali month names (transliterated)
NEPALI_MONTHS = [
    "Baisakh",
    "Jestha",
    "Ashadh",
    "Shrawan",
    "Bhadra",
    "Ashwin",
    "Kartik",
    "Mangsir",
    "Poush",
    "Magh",
    "Falgun",
    "Chaitra",
]

# DARK MODE COLORS
BG = "#0F172A"  # main background
CARD = "#1E293B"  # card background
ACCENT = "#3B82F6"  # primary blue
ACCENT_DARK = "#1D4ED8"
MUTED = "#94A3B8"  # subtle text
TEXT = "#E2E8F0"  # main text
BORDER = "#334155"


class TenantManagerGUI:

    def load_settings(self):
        try:
            s = pd.read_excel("settings.xlsx")
            self.settings = s.set_index("Key")["Value"].to_dict()
        except:
            self.settings = {"rent": 4000}

        self.monthly_rent = int(self.settings.get("rent", 4000))

    def calculate_totals(self, idx):
        row = self.df.loc[idx]

        total = (
            self.safe_int(row["Rent"])
            + self.safe_int(row["Wtr Chg"])
            + self.safe_int(row["Elec Chg"])
            + self.safe_int(row["Net Chg"])
            + self.safe_int(row["Misc"])
            + self.safe_int(row["Balance"])
        )

        paid = self.safe_int(row["Paid"])
        base_advance = self.safe_int(row.get("Advance", 0))

        # ✅ STEP 1: Combine available money
        effective_paid = paid + base_advance

        # ✅ STEP 2: Calculate outstanding
        outstanding = max(total - effective_paid, 0)

        # ✅ STEP 3: Calculate new advance (wallet logic)
        # What remains after paying the bill
        new_advance = max(effective_paid - total, 0)

        return total, outstanding, new_advance

    def get_base_advance(self, rid, idx):
        if rid not in self._base_advance_map:
            self._base_advance_map[rid] = self.safe_int(self.df.at[idx, "Advance"])
        return self._base_advance_map[rid]

    def _save_df_async(self):
        self.root.after(10, self._save_df)

    def safe_int(self, val, minimum=0):
        try:
            return max(int(val), minimum)
        except:
            return minimum

    def start_edit_cell(self, event, tree, record_id, year, month):
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        row_id = tree.identify_row(event.y)
        col = tree.identify_column(event.x)

        if not row_id:
            return

        col_index = int(col.replace("#", "")) - 1
        values = list(tree.item(row_id, "values"))

        # Only allow editing certain columns (Units, Rate, Charge)

        item_name = values[0]

        if not (
            (item_name in ["Water", "Electricity", "Internet"] and col_index in [1, 2])
            or (item_name == "Misc" and col_index == 3)
            or (item_name == "Paid" and col_index == 3)
        ):
            return

        # Get cell position

        bbox = tree.bbox(row_id, col)
        if not bbox:
            return

        x, y, width, height = bbox

        # Create entry box
        entry = tk.Entry(tree)

        def only_numbers(event):
            val = entry.get()
            if val == "":
                return
            cleaned = "".join(filter(str.isdigit, val))
            if val != cleaned:
                pos = entry.index(tk.INSERT)
                entry.delete(0, tk.END)
                entry.insert(0, cleaned)
                entry.icursor(max(0, pos - 1))

        entry.bind("<KeyRelease>", only_numbers)

        entry.place(x=x, y=y, width=width, height=height)

        # Pre-fill value

        initial = values[col_index]
        if initial == "-":
            initial = "0"

        entry.insert(0, initial)

        entry.focus()
        entry.select_range(0, tk.END)

        def save_edit(event=None):
            try:
                new_val = entry.get().strip()

                if new_val == "":
                    new_val = "0"

                new_val = self.safe_int(new_val)
                values[col_index] = new_val

                # ✅ Auto recalc for utilities
                item_name = values[0]

                if item_name in ["Water", "Electricity", "Internet"]:

                    units = self.safe_int(values[1])
                    rate = self.safe_int(values[2])

                    # If editing charge directly
                    if col_index == 3:
                        values[1] = 0
                        values[2] = 0
                    else:
                        values[3] = units * rate

                # Update UI
                tree.item(row_id, values=values)
                entry.destroy()

                # ✅ Update DataFrame
                idx = self._id_index.get(str(record_id))
                if idx is None:
                    return

                if item_name == "Water":
                    self.df.at[idx, "Wtr Units"] = self.safe_int(values[1])
                    self.df.at[idx, "Wtr Rate"] = self.safe_int(values[2])

                    # ✅ ALWAYS recalc from units x rate
                    self.df.at[idx, "Wtr Chg"] = (
                        self.df.at[idx, "Wtr Units"] * self.df.at[idx, "Wtr Rate"]
                    )

                elif item_name == "Electricity":
                    self.df.at[idx, "Elec Units"] = self.safe_int(values[1])
                    self.df.at[idx, "Elec Rate"] = self.safe_int(values[2])
                    self.df.at[idx, "Elec Chg"] = (
                        self.df.at[idx, "Elec Units"] * self.df.at[idx, "Elec Rate"]
                    )

                elif item_name == "Internet":
                    self.df.at[idx, "Net Units"] = self.safe_int(values[1])
                    self.df.at[idx, "Net Rate"] = self.safe_int(values[2])
                    self.df.at[idx, "Net Chg"] = (
                        self.df.at[idx, "Net Units"] * self.df.at[idx, "Net Rate"]
                    )

                elif item_name == "Misc":
                    self.df.at[idx, "Misc"] = self.safe_int(values[3])

                elif item_name == "Paid":
                    self.df.at[idx, "Paid"] = self.safe_int(values[3])

                rid = str(self.df.at[idx, "RecordID"])

                base_advance = self.get_base_advance(rid, idx)
                self.df.at[idx, "Advance"] = base_advance

                total, outstanding, new_advance = self.calculate_totals(idx)

                self.df.at[idx, "Total"] = total
                self.df.at[idx, "Outstanding"] = outstanding
                self.df.at[idx, "Advance"] = new_advance

                self.schedule_save()
                self._last_data_hash[str(record_id)] = None
                self.refresh_current_frame()

            except:
                entry.destroy()

        # Save on Enter

        original_value = values[col_index]  # ✅ store original for ESC

        def move(row, col_id):
            bbox = tree.bbox(row, col_id)
            if not bbox:
                return
            event.x = bbox[0] + 5
            event.y = bbox[1] + 5
            self.start_edit_cell(event, tree, record_id, year, month)

        # ✅ ENTER → Save only
        def on_enter(e):
            save_edit()
            return "break"

        # ✅ ESC → cancel edit (restore old value)

        def on_escape(e):
            # ✅ restore original value
            values[col_index] = original_value
            tree.item(row_id, values=values)

            entry.destroy()
            return "break"

        # ✅ TAB → next column or next row
        def on_tab(e):
            save_edit()

            if col_index < 3:
                next_col = f"#{col_index + 2}"
                move(row_id, next_col)
            else:
                next_row = tree.next(row_id)
                if next_row:
                    move(next_row, "#2")  # jump to Units
            return "break"

        # ✅ SHIFT + TAB → previous column or prev row
        def on_shift_tab(e):
            save_edit()

            if col_index > 1:
                prev_col = f"#{col_index}"
                move(row_id, prev_col)
            else:
                prev_row = tree.prev(row_id)
                if prev_row:
                    move(prev_row, "#4")  # go to Charge
            return "break"

        entry.bind("<Return>", on_enter)
        entry.bind("<Escape>", on_escape)
        entry.bind("<Tab>", on_tab)
        entry.bind("<Shift-Tab>", on_shift_tab)

    def open_settings(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Settings")
        dlg.geometry("300x200")
        dlg.configure(bg="#0F172A")

        tk.Label(dlg, text="Monthly Rent", bg="#0F172A", fg="white").pack(pady=10)

        rent_entry = ttk.Entry(dlg)
        rent_entry.pack(padx=20)
        rent_entry.insert(0, str(self.monthly_rent))

        def save():
            try:
                new_rent = int(rent_entry.get())

                # ✅ update settings
                self.settings["rent"] = new_rent

                df = pd.DataFrame(self.settings.items(), columns=["Key", "Value"])
                df.to_excel("settings.xlsx", index=False)

                # ✅ apply only for FUTURE
                self.monthly_rent = new_rent

                messagebox.showinfo(
                    "Saved", "Rent updated ✅\nOnly affects FUTURE records."
                )
                dlg.destroy()

            except:
                messagebox.showerror("Error", "Invalid rent value")

        ttk.Button(dlg, text="Save", command=save).pack(pady=20)

    def toggle_select_all(self):
        current = getattr(self, "_last_selected_frame_key", None)

        if not current:
            return

        cur_year, cur_month = current

        # ✅ FIND tenants only in current month
        current_ids = []

        for rid, data in self._ui_map.items():
            if rid == "_summary":
                continue

            idx = self._id_index.get(rid)
            if idx is None:
                continue

            row = self.df.iloc[idx]

            if str(row["Year"]) == str(cur_year) and str(row["Month"]) == str(
                cur_month
            ):
                current_ids.append(rid)

        # ✅ check if all of them are already selected
        all_selected = all(rid in self.selected_ids for rid in current_ids)

        # ✅ toggle ONLY current month tenants
        for rid in current_ids:

            data = self._ui_map.get(rid)
            if not data:
                continue

            var = data.get("selected")
            if not var:
                continue

            var.set(not all_selected)

            if not all_selected:
                self.selected_ids.add(rid)
            else:
                self.selected_ids.discard(rid)

    def clean_credit(self, val):
        if pd.isna(val):
            return ""
        val = str(val).strip()
        return "" if val.lower() in ["nan", "none"] else val

    def get_row_color(self, row, dim=False):
        try:
            out = self.safe_int(row["Outstanding"])
            adv = self.safe_int(row["Advance"])

            if out > 0:
                return "#3B1F1F"  # red
            elif adv > 0:
                return "#1E3A2F"  # green
            else:
                return "#111827" if dim else CARD
        except:
            return CARD

    def _build_index(self):
        self._id_index = {rid: i for i, rid in enumerate(self.df["RecordID"])}

    def _build_month_cache(self):
        self._month_cache = {}

        for _, row in self.df.iterrows():
            key = (str(row["Year"]), str(row["Month"]))
            self._month_cache.setdefault(key, []).append(row)

    def carry_over_by_id(self, record_id):
        idx = self._id_index.get(str(record_id))
        if idx is None:
            return

        row = self.df.iloc[idx]
        self.carry_over_tenant_dialog(row)

    def delete_by_id(self, record_id):
        idx = self._id_index.get(str(record_id))
        if idx is None:
            return

        row = self.df.iloc[idx]
        self.delete_tenant(row)

    def open_edit_by_id(self, rid):
        match = self.df[self.df["RecordID"] == rid]

        if match.empty:
            messagebox.showerror("Error", "Record no longer exists.")
            return

        row = match.iloc[0]
        self.edit_full_record_dialog(row)

    def current_time_str(self):
        return datetime.now().strftime("%I:%M %p")  # e.g., "02:45 PM"

    def set_sync_status(self, text, color="#22C55E"):
        try:
            if hasattr(self, "sync_status"):

                # ✅ Add time when status is "Synced"
                if "✅ Synced" in text:
                    text = f"✅ Synced • {self.current_time_str()}"

                self.sync_status.config(text=text, fg=color)
                self.root.update_idletasks()

            # ✅ cancel previous timer
            if self._sync_timer:
                try:
                    self.root.after_cancel(self._sync_timer)
                except:
                    pass

            # ✅ auto return to synced after loading
            if "⏳" in text or "🔄" in text:
                self._sync_timer = self.root.after(
                    1200, lambda: self.set_sync_status("✅ Synced", "#22C55E")
                )

        except:
            pass

    def remove_month_from_sidebar(self, year, month):
        if not hasattr(self, "sidebar_tree"):
            return

        tree = self.sidebar_tree

        target_iid = f"ym_{year}_{month or 'Unspecified'}"
        parent_iid = f"year_{year}"

        # ✅ remove month node
        try:
            if tree.exists(target_iid):
                tree.delete(target_iid)
        except:
            pass

        # ✅ if no months left → remove year node
        try:
            if tree.exists(parent_iid):
                children = tree.get_children(parent_iid)
                if not children:
                    tree.delete(parent_iid)
        except:
            pass

    def build_month_ui(self, frame, year, month):

        month_data = self.df[
            (self.df["Year"].astype(str) == str(year))
            & (self.df["Month"].astype(str) == str(month))
        ].copy()

        month_data = month_data.sort_values(by=["Block", "Room"])

        row_idx, col_idx = 0, 0

        for _, row in month_data.iterrows():

            tenant_frame = tk.Frame(
                frame, bg=CARD, bd=1, relief="solid", highlightbackground=BORDER
            )

            self.set_frame_color(tenant_frame, self.get_row_color(row))

            tenant_frame.grid(
                row=row_idx, column=col_idx, padx=10, pady=10, sticky="nsew"
            )
            tenant_frame.grid_propagate(False)
            tenant_frame.config(width=360)

            # ✅ checkbox
            rid = str(row["RecordID"])
            var = tk.BooleanVar(value=(rid in self.selected_ids))

            def on_check(*args):
                if var.get():
                    self.selected_ids.add(rid)
                else:
                    self.selected_ids.discard(rid)

            var.trace_add("write", on_check)

            cb = ttk.Checkbutton(tenant_frame, variable=var)
            cb.pack(anchor="ne")

            # ✅ header
            header = tk.Label(
                tenant_frame,
                text=f"🏢 Block {row['Block']}    🚪 Room {row['Room']}",
                bg=tenant_frame.cget("bg"),
                fg="#FACC15",
                font=("Segoe UI", 10, "bold"),
            )
            header.pack(anchor="nw", padx=8, pady=(8, 0))

            tenant_lbl = tk.Label(
                tenant_frame,
                text=row["Name"],
                bg=tenant_frame.cget("bg"),
                fg="#93C5FD",
                font=("Segoe UI", 10),
            )
            tenant_lbl.pack(anchor="nw", padx=8, pady=(0, 4))
            tenant_lbl.bind(
                "<Double-Button-1>",
                lambda e, rid=row["RecordID"]: self.open_edit_by_id(rid),
            )

            # ✅ Table
            cols = ["Item", "Units", "Rate", "Charge"]
            tree = ttk.Treeview(tenant_frame, columns=cols, show="headings", height=10)

            for c in cols:
                tree.heading(c, text=c)

            tree.column("Item", width=120)
            tree.column("Units", width=60)
            tree.column("Rate", width=60)
            tree.column("Charge", width=90, anchor="e")

            def add(iid, values):
                tree.insert("", "end", iid=iid, values=values)

            add(rid + "_rent", ("Rent", "-", "-", self.safe_int(row["Rent"])))
            add(
                rid + "_elec",
                (
                    "Electricity",
                    self.safe_int(row["Elec Units"]),
                    self.safe_int(row["Elec Rate"]),
                    self.safe_int(row["Elec Chg"]),
                ),
            )
            add(
                rid + "_net",
                (
                    "Internet",
                    self.safe_int(row["Net Units"]),
                    self.safe_int(row["Net Rate"]),
                    self.safe_int(row["Net Chg"]),
                ),
            )
            add(
                rid + "_water",
                (
                    "Water",
                    self.safe_int(row["Wtr Units"]),
                    self.safe_int(row["Wtr Rate"]),
                    self.safe_int(row["Wtr Chg"]),
                ),
            )
            add(rid + "_misc", ("Misc", "-", "-", self.safe_int(row["Misc"])))
            add(rid + "_bal", ("Old Balance", "-", "-", self.safe_int(row["Balance"])))

            advance_val = row["Advance"]

            try:
                advance_val = int(advance_val)
            except:
                advance_val = 0

            add(rid + "_advance", ("Advance (Auto)", "-", "-", advance_val))

            add(rid + "_total", ("Total Bill", "-", "-", self.safe_int(row["Total"])))
            add(
                rid + "_paid",
                (
                    "Paid",
                    self.clean_credit(row.get("Credit", "")) or "-",
                    "-",
                    self.safe_int(row["Paid"]),
                ),
            )
            add(
                rid + "_outstanding",
                ("Outstanding", "-", "-", self.safe_int(row["Outstanding"])),
            )

            tree.pack(fill="x", padx=8, pady=8)

            tree.bind(
                "<Double-1>",
                lambda e, t=tree, rid=row["RecordID"], y=row["Year"], m=row[
                    "Month"
                ]: self.start_edit_cell(e, t, rid, y, m),
            )

            # ✅ Notes
            notes = str(row["Notes"]).strip()

            note_lbl = tk.Label(
                tenant_frame,
                text=f"📝 {notes or '(No notes)'}",
                bg=tenant_frame.cget("bg"),
                fg="#CBD5F5" if notes else "#64748B",
                cursor="hand2",
                wraplength=320,
            )
            note_lbl.pack(anchor="w", padx=10, pady=(0, 6))

            note_lbl.bind(
                "<Button-1>", lambda e, rid=row["RecordID"]: self.edit_notes(rid)
            )

            # ✅ Buttons
            btns = tk.Frame(tenant_frame, bg=tenant_frame.cget("bg"))
            btns.pack(pady=(0, 8))

            ttk.Button(
                btns,
                text="Edit",
                command=lambda rid=row["RecordID"]: self.open_edit_by_id(rid),
            ).pack(side="left", padx=4)

            ttk.Button(
                btns,
                text="Carry",
                command=lambda rid=row["RecordID"]: self.carry_over_by_id(rid),
            ).pack(side="left", padx=4)

            ttk.Button(
                btns,
                text="Delete",
                command=lambda rid=row["RecordID"]: self.delete_by_id(rid),
            ).pack(side="left", padx=4)

            # ✅ total label
            total_lbl = tk.Label(
                tenant_frame,
                text=f"Tenant Total: {self.safe_int(row['Total'])}",
                bg=tenant_frame.cget("bg"),
                fg="#60A5FA",
                font=("Segoe UI", 9, "bold"),
            )
            total_lbl.pack(anchor="e", padx=8, pady=(0, 6))

            # ✅ IMPORTANT: restore _ui_map for search/edit/refresh
            self._ui_map[rid] = {
                "tree": tree,
                "total_label": total_lbl,
                "selected": var,
                "frame": tenant_frame,
                "header_label": header,
                "tenant_label": tenant_lbl,
                "notes_label": note_lbl,  # ✅ ADD THIS
            }

            col_idx += 1
            if col_idx == 4:
                col_idx = 0
                row_idx += 1

        # ✅ Summary
        summary = tk.Label(
            frame,
            text=f"Total Collected: {self.safe_int(month_data['Paid'].sum())}",
            bg=CARD,
            fg="#22C55E",
            font=("Segoe UI", 10, "bold"),
        )
        summary.grid(row=row_idx + 1, column=0, columnspan=3, pady=10)

        self._ui_map.setdefault("_summary", {})[(str(year), str(month))] = summary

    def edit_notes(self, record_id):
        idx = self._id_index.get(str(record_id))
        if idx is None:
            return

        current = str(self.df.at[idx, "Notes"])

        new = simpledialog.askstring(
            "Edit Notes",
            "Update notes:",
            initialvalue=current,
            parent=self.show_all_window,
        )

        if new is None:
            return

        self.df.at[idx, "Notes"] = new
        self._save_df()

    def _on_date_search(self, *args):
        if not self.show_all_window:
            return

        self._filter_sidebar_only()

    def _on_tenant_search(self, *args):
        if self._search_after_id:
            self.root.after_cancel(self._search_after_id)

        self._search_after_id = self.root.after(80, self._apply_tenant_search)

    def _apply_tenant_search(self):
        if not self._ui_map:
            return

        query_raw = self.tenant_search_var.get().strip()
        query = query_raw.lower()

        if query == "search tenant...":
            query = ""

        first_match_frame = None

        for key, item in self._ui_map.items():
            if key == "_summary":
                continue

            rid = key

            idx = self._id_index.get(rid)
            if idx is None:
                continue

            row = self.df.iloc[idx]

            frame = item["frame"]
            header = item["header_label"]
            tenant_lbl = item["tenant_label"]

            name = str(row["Name"])
            text = f"{name} {row['Block']} {row['Room']}".lower()

            # ✅ RESET (no search)
            if query == "" or query == "search tenant...":

                out = self.safe_int(row["Outstanding"])
                adv = self.safe_int(row["Advance"])

                if out > 0:
                    self.set_frame_color(frame, "#3B1F1F")
                elif adv > 0:
                    self.set_frame_color(frame, "#1E3A2F")
                else:
                    self.set_frame_color(frame, CARD)

                frame.config(highlightthickness=0)

                header.config(
                    text=f"🏢 Block {row['Block']}    🚪 Room {row['Room']}",
                    fg="#FACC15",
                )

                tenant_lbl.config(
                    text=name,
                    fg="#93C5FD",
                )
                continue

            # ✅ MATCH CASE
            if query in text:

                frame.config(highlightbackground="#3B82F6", highlightthickness=2)
                self.set_frame_color(frame, "#1E293B")

                header.config(
                    text=f"🏢 Block {row['Block']}    🚪 Room {row['Room']}",
                    fg="#FACC15",
                )

                # ✅ highlight name
                lower_name = name.lower()
                idx = lower_name.find(query)

                if idx != -1:
                    end = idx + len(query)
                    highlighted = name[:idx] + "🔎" + name[idx:end] + "🔎" + name[end:]
                    tenant_lbl.config(text=highlighted, fg="#FACC15")
                else:
                    tenant_lbl.config(text=name, fg="#93C5FD")

                if first_match_frame is None:
                    first_match_frame = frame

            # ✅ NON-MATCH CASE
            else:

                frame.config(highlightthickness=0)

                self.set_frame_color(frame, self.get_row_color(row, dim=True))

                header.config(
                    text=f"🏢 Block {row['Block']}    🚪 Room {row['Room']}",
                    fg="#64748B",
                )

                tenant_lbl.config(
                    text=name,
                    fg="#64748B",
                )

        # ✅ AUTO SCROLL
        if first_match_frame and hasattr(self, "canvas"):
            try:
                self.canvas.update_idletasks()
                y = first_match_frame.winfo_rooty() - self.canvas.winfo_rooty()
                height = max(1, first_match_frame.master.winfo_height())
                self.canvas.yview_moveto(max(0, min(1, y / height)))
            except:
                pass

    def edit_full_record_dialog(self, row):
        dlg = tk.Toplevel(self.show_all_window or self.root)
        dlg.title("Edit Tenant")
        dlg.geometry("380x580")

        dlg.update_idletasks()
        x = (dlg.winfo_screenwidth() // 2) - (380 // 2)
        y = (dlg.winfo_screenheight() // 2) - (580 // 2)
        dlg.geometry(f"+{x}+{y}")

        dlg.transient(self.show_all_window or self.root)
        dlg.grab_set()
        dlg.configure(bg="#0F172A")

        widgets = {}

        # ✅ Scrollable scroll_frame setup
        canvas = tk.Canvas(dlg, bg="#0F172A", highlightthickness=0)
        scrollbar = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)

        scroll_frame = tk.Frame(canvas, bg="#0F172A")

        scroll_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def resize_frame(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind("<Configure>", resize_frame)

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.configure(yscrollincrement=20)
        canvas.configure(confine=True)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # ✅ Mouse wheel scrolling for dialog
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind(
            "<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel)
        )
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # ✅ Card section

        def section(title):
            card = tk.Frame(scroll_frame, bg="#1E293B", bd=1, relief="solid")
            card.pack(fill="x", pady=8)

            header = tk.Label(
                card,
                text=title,
                bg="#1E293B",
                fg="#60A5FA",
                font=("Segoe UI", 11, "bold"),
            )
            header.pack(anchor="w", padx=10, pady=(8, 4))

            body = tk.Frame(card, bg="#1E293B")
            body.pack(fill="x", padx=10, pady=(0, 10))

            return body

        def row_field(parent, label, field):
            rowf = tk.Frame(parent, bg="#1E293B")
            rowf.pack(fill="x", pady=4)

            tk.Label(
                rowf,
                text=label,
                bg="#1E293B",
                fg="#E2E8F0",
                width=16,
                anchor="w",
                font=("Segoe UI", 9),
            ).pack(side="left")

            raw = row.get(field, "")

            if pd.isna(raw) or str(raw).lower() == "nan":
                val = ""
            else:
                val = str(raw)

            if field == "Year":
                e = ttk.Combobox(rowf, values=self._year_range(), width=20)
                e.set(val or str(current_nepali_year()))

            elif field == "Month":
                e = ttk.Combobox(rowf, values=NEPALI_MONTHS, width=20)
                e.set(val or NEPALI_MONTHS[0])

            else:
                e = ttk.Entry(rowf, width=26)
                e.insert(0, val)

                if field == "Advance":
                    e.config(state="readonly")

            e.pack(side="right", fill="x", expand=True)
            widgets[field] = e

            # ✅ allow only numbers but not strict (auto clean)
            def clean_numeric(event):
                val = e.get()

                if val == "":
                    return

                cleaned = "".join(filter(str.isdigit, val))

                if val != cleaned:
                    pos = e.index(tk.INSERT)
                    e.delete(0, tk.END)
                    e.insert(0, cleaned)
                    e.icursor(min(pos - 1, len(cleaned)))

            # ✅ apply only to numeric fields

            if field not in ["Name", "Block", "Room", "Month", "Year", "Credit"]:

                def on_key(event):
                    clean_numeric(event)
                    update_preview()

                e.bind("<KeyRelease>", on_key)

        # ✅ Sections
        basic = section("👤 Basic Info")
        row_field(basic, "Name", "Name")
        row_field(basic, "Block", "Block")
        row_field(basic, "Room", "Room")

        date = section("📅 Date")
        row_field(date, "Year", "Year")
        row_field(date, "Month", "Month")

        util = section("💡 Utilities")
        row_field(util, "Water Units", "Wtr Units")
        row_field(util, "Water Rate", "Wtr Rate")
        row_field(util, "Water Charge", "Wtr Chg")
        row_field(util, "Electric Units", "Elec Units")
        row_field(util, "Electric Rate", "Elec Rate")
        row_field(util, "Electric Charge", "Elec Chg")
        row_field(util, "Net Units", "Net Units")
        row_field(util, "Net Rate", "Net Rate")
        row_field(util, "Net Charge", "Net Chg")

        finan = section("💰 Financial")
        row_field(finan, "Misc", "Misc")
        row_field(finan, "Advance", "Advance")
        row_field(finan, "Balance", "Balance")
        row_field(finan, "Paid", "Paid")
        row_field(finan, "Credit", "Credit")

        def setup_manual_logic(units_f, rate_f, charge_f):
            u = widgets[units_f]
            r = widgets[rate_f]
            c = widgets[charge_f]

            def on_charge_change(event):
                val = c.get().strip()

                try:
                    val_int = int(val)
                except:
                    val_int = 0

                if val_int > 0:

                    # ✅ manual mode
                    u.delete(0, tk.END)
                    r.delete(0, tk.END)

                    u.insert(0, "0")
                    r.insert(0, "0")

                    u.config(state="disabled")
                    r.config(state="disabled")
                else:
                    # ✅ auto mode
                    u.config(state="normal")
                    r.config(state="normal")

            def on_units_rate_change(event):
                # ✅ if user edits units/rate → clear manual charge
                if c.get().strip():
                    c.delete(0, tk.END)
                    u.config(state="normal")
                    r.config(state="normal")

            c.bind("<KeyRelease>", on_charge_change)
            u.bind("<KeyRelease>", on_units_rate_change)
            r.bind("<KeyRelease>", on_units_rate_change)

        setup_manual_logic("Wtr Units", "Wtr Rate", "Wtr Chg")
        setup_manual_logic("Elec Units", "Elec Rate", "Elec Chg")
        setup_manual_logic("Net Units", "Net Rate", "Net Chg")

        def apply_initial_state(units_f, rate_f, charge_f):
            c = widgets[charge_f]
            u = widgets[units_f]
            r = widgets[rate_f]

            val = c.get().strip()

            try:
                val_int = int(val)
            except:
                val_int = 0

            if val_int > 0:
                u.config(state="disabled")
                r.config(state="disabled")

        apply_initial_state("Wtr Units", "Wtr Rate", "Wtr Chg")
        apply_initial_state("Elec Units", "Elec Rate", "Elec Chg")
        apply_initial_state("Net Units", "Net Rate", "Net Chg")

        widgets["Name"].focus()
        widgets["Name"].select_range(0, tk.END)

        # ✅ Notes
        notes_sec = section("📝 Notes")
        notes_box = tk.Text(notes_sec, height=4)
        notes_box.pack(fill="x")
        notes_box.insert("1.0", str(row.get("Notes", "")))

        # ✅ Preview card
        preview_card = section("📊 Summary")

        preview_total = tk.Label(
            preview_card,
            text="Total: 0",
            fg="#38BDF8",
            bg="#1E293B",
            font=("Segoe UI", 11, "bold"),
        )
        preview_total.pack(anchor="w", pady=2)

        preview_out = tk.Label(
            preview_card,
            text="Outstanding: 0",
            fg="#F87171",
            bg="#1E293B",
            font=("Segoe UI", 11, "bold"),
        )
        preview_out.pack(anchor="w", pady=2)

        # ✅ Live calc
        def update_preview():
            try:

                def get_int(f):
                    try:
                        val = widgets[f].get().strip()
                        return max(int(val), 0) if val != "" else 0
                    except:
                        return 0

                def calc(units_f, rate_f, charge_f):
                    charge = get_int(charge_f)
                    units = get_int(units_f)
                    rate = get_int(rate_f)

                    # ✅ If manual charge entered → use it
                    if charge > 0:
                        return charge

                    # ✅ else calculate
                    return units * rate

                wtr = calc("Wtr Units", "Wtr Rate", "Wtr Chg")
                elec = calc("Elec Units", "Elec Rate", "Elec Chg")
                net = calc("Net Units", "Net Rate", "Net Chg")

                total = (
                    self.safe_int(row["Rent"])
                    + wtr
                    + elec
                    + net
                    + get_int("Misc")
                    + get_int("Balance")
                )

                paid = get_int("Paid")

                if paid >= total:
                    outstanding = 0
                else:
                    outstanding = total - paid

                preview_total.config(text=f"Total: {total}")
                preview_out.config(
                    text=f"Outstanding: {outstanding}",
                    fg="#F87171" if outstanding > 0 else "#22C55E",
                )
            except:
                pass

        update_preview()

        # ✅ SAVE
        def save():
            try:
                rid = str(row["RecordID"])

                idx = self._id_index.get(rid)
                if idx is None:
                    messagebox.showerror("Error", "Record not found")
                    return

                for f in widgets:
                    val = widgets[f].get().strip()

                    if f == "Credit":
                        self.df.at[idx, "Credit"] = val.strip() if val else ""
                    elif f in ["Name", "Block", "Room", "Month", "Year"]:
                        self.df.at[idx, f] = val

                    else:
                        try:
                            self.df.at[idx, f] = max(int(val or 0), 0)
                        except:
                            self.df.at[idx, f] = 0  # ✅ safe fallback

                # ✅ FORCE numeric safety (this fixes crash)
                for col in [
                    "Wtr Units",
                    "Wtr Rate",
                    "Elec Units",
                    "Elec Rate",
                    "Net Units",
                    "Net Rate",
                    "Misc",
                    "Balance",
                    "Paid",
                ]:
                    try:
                        self.df.at[idx, col] = int(self.df.at[idx, col])
                    except:
                        self.df.at[idx, col] = 0

                self.df.at[idx, "Notes"] = notes_box.get("1.0", tk.END).strip()

                def calc_save(units, rate, charge):
                    charge = int(charge)
                    if charge > 0:
                        return charge
                    return int(units) * int(rate)

                wtr = calc_save(
                    self.df.at[idx, "Wtr Units"],
                    self.df.at[idx, "Wtr Rate"],
                    self.df.at[idx, "Wtr Chg"],
                )

                elec = calc_save(
                    self.df.at[idx, "Elec Units"],
                    self.df.at[idx, "Elec Rate"],
                    self.df.at[idx, "Elec Chg"],
                )

                net = calc_save(
                    self.df.at[idx, "Net Units"],
                    self.df.at[idx, "Net Rate"],
                    self.df.at[idx, "Net Chg"],
                )

                self.df.at[idx, "Wtr Chg"] = wtr
                self.df.at[idx, "Elec Chg"] = elec
                self.df.at[idx, "Net Chg"] = net

                total = (
                    self.df.at[idx, "Rent"]
                    + wtr
                    + elec
                    + net
                    + self.df.at[idx, "Misc"]
                    + self.df.at[idx, "Balance"]
                )
                paid = int(self.df.at[idx, "Paid"])
                credit = str(self.df.at[idx, "Credit"]).strip()

                # ✅ APPLY CREDIT (optional but correct behavior)
                if credit:
                    # just store, don't convert — display only
                    pass

                self.df.at[idx, "Total"] = total

                rid = str(self.df.at[idx, "RecordID"])

                self._base_advance_map[rid] = self.safe_int(self.df.at[idx, "Advance"])

                base_advance = self.get_base_advance(rid, idx)
                self.df.at[idx, "Advance"] = base_advance

                total, outstanding, new_advance = self.calculate_totals(idx)

                self.df.at[idx, "Total"] = total
                self.df.at[idx, "Outstanding"] = outstanding
                self.df.at[idx, "Advance"] = new_advance

                self._save_df()
                dlg.destroy()
                self.refresh_current_frame()

            except Exception as e:
                messagebox.showerror("Error", f"Invalid input:\n{e}")

        dlg.bind("<Return>", lambda e: save())

        # ✅ Buttons
        btns = tk.Frame(scroll_frame, bg="#0F172A")
        btns.pack(pady=12)

        ttk.Button(btns, text="✅ Save", command=save).pack(side="left", padx=10)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=10)

    def _watch_excel_changes(self):
        try:
            if os.path.exists(self.filename):
                current_mtime = os.path.getmtime(self.filename)

                if not hasattr(self, "_last_mtime"):
                    self._last_mtime = current_mtime

                # ✅ if file changed
                elif current_mtime != self._last_mtime:
                    self._last_mtime = current_mtime

                    # reload safely
                    self.set_sync_status("🔄 Reloading...", "#38BDF8")
                    self._load_or_create_df(self.filename)
                    self._build_index()
                    self._build_month_cache()

                    # refresh UI if open
                    if self.show_all_window and tk.Toplevel.winfo_exists(
                        self.show_all_window
                    ):
                        self.refresh_current_frame()
                        self.set_sync_status("✅ Synced", "#22C55E")

        except Exception:
            pass  # silent fail (important)

        # ✅ run every 10 seconds

        if not self.root.winfo_exists():
            return

        self.root.after(15000, self._watch_excel_changes)

    def _filter_sidebar_only(self):
        if not hasattr(self, "sidebar_tree"):
            return

        q = self.date_search_var.get().strip().lower()
        tree = self.sidebar_tree

        for node in tree.get_children():
            tree.delete(node)

        # rebuild only sidebar (NOT UI)
        self.df["Year"] = self.df["Year"].astype(str)

        valid_df = self.df[self.df["Name"].astype(str).str.strip() != ""]

        combos = (
            valid_df[["Year", "Month"]].drop_duplicates().fillna("").values.tolist()
        )

        year_map = {}
        for y, m in combos:
            year_map.setdefault(str(y), []).append(str(m))

        for year, months in year_map.items():
            if q and q not in year.lower():
                continue

            year_node = tree.insert("", "end", text=year, open=True)

            for m in months:
                display = m or "Unspecified"
                if not q or q in display.lower():
                    tree.insert(year_node, "end", text=display)

    def set_frame_color(self, frame, color):
        frame.config(bg=color)
        for child in frame.winfo_children():
            try:
                child.config(bg=color)
            except:
                pass

    def schedule_save(self):
        if self._save_after_id:
            self.root.after_cancel(self._save_after_id)

        # ✅ save after user stops editing (1.5 sec)
        self._save_after_id = self.root.after(5000, self._save_df)

    def __init__(self, root, filename="tenants.xlsx"):
        self.entries = {}
        self._pending_navigation = None
        self._is_rebuilding_ui = False
        self._sync_timer = None
        self.load_settings()
        self.selected_ids = set()
        self.root = root
        self.filename = filename
        self.show_all_window = None
        self._save_after_id = None
        self._base_advance_map = {}
        self._last_data_hash = {}
        # UI maps
        self._ui_map = {}
        self._frame_map = {}

        self.tenant_search_var = tk.StringVar()
        self.date_search_var = tk.StringVar()

        self.date_search_var.trace_add("write", self._on_date_search)
        self.tenant_search_var.trace_add("write", self._on_tenant_search)

        self._search_after_id = None

        # Apply global styling
        self._apply_style()

        # Load or create Excel file
        self._load_or_create_df(filename)
        self._build_index()
        self._build_month_cache()

        # Top header
        header = tk.Frame(root, bg=ACCENT, height=64)
        header.pack(fill="x")
        header.pack_propagate(False)
        title_font = tkfont.Font(family="Segoe UI", size=16, weight="bold")
        tk.Label(
            header, text="Tenant Manager", bg=ACCENT, fg="white", font=title_font
        ).pack(side="left", padx=16)
        tk.Label(
            header,
            text="— Clean • Fast • Nepali months",
            bg=ACCENT,
            fg=MUTED,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=8)

        # Main content frame with subtle background
        main = tk.Frame(root, bg=BG)
        main.pack(fill="both", expand=True)

        # Left column: Add Tenant card
        left_col = tk.Frame(main, bg=BG)
        left_col.pack(side="left", fill="y", padx=16, pady=16)

        card = tk.Frame(
            left_col, bg=CARD, bd=1, relief="solid", highlightbackground="#475569"
        )
        card.pack(fill="y", padx=4, pady=4)
        card.grid_propagate(False)
        card.config(width=360)

        card_title = tk.Label(
            card,
            text="Tenant Actions",
            bg=CARD,
            fg=TEXT,
            font=("Segoe UI", 12, "bold"),
        )
        card_title.grid(
            row=0, column=0, columnspan=2, sticky="w", padx=12, pady=(12, 6)
        )

        # Action buttons
        btn_frame = tk.Frame(card, bg=CARD)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=12)

        add_btn = ttk.Button(
            btn_frame,
            text="Add Tenant",
            command=self.add_new_dialog,
            style="Accent.TButton",
        )
        add_btn.pack(side="left", padx=8)
        show_btn = ttk.Button(btn_frame, text="Open Dashboard", command=self.show_all)
        show_btn.pack(side="left", padx=8)

        # Right column: welcome card with big Open Dashboard button (in case auto-open disabled)
        right_col = tk.Frame(main, bg=BG)
        right_col.pack(side="left", fill="both", expand=True, padx=(0, 16), pady=16)

        welcome_card = tk.Frame(right_col, bg=CARD)
        welcome_card.pack(fill="both", expand=True)
        tk.Label(
            welcome_card,
            text="Overview",
            bg=CARD,
            fg=TEXT,
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="nw", padx=12, pady=(12, 6))
        tk.Label(
            welcome_card,
            text="Click the button below to open the interactive dashboard.",
            bg=CARD,
            fg=MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="nw", padx=12, pady=(0, 12))

        open_btn_frame = tk.Frame(welcome_card, bg=CARD)
        open_btn_frame.pack(expand=True, fill="both", pady=20)
        open_btn = ttk.Button(
            open_btn_frame,
            text="Open Dashboard",
            command=self.show_all,
            style="Accent.TButton",
        )
        open_btn.place(relx=0.5, rely=0.5, anchor="center")

        # Keep root minimum size
        root.minsize(1000, 600)

        # Start watching for external Excel changes
        self._watch_excel_changes()

    def _apply_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        default_font = ("Segoe UI", 10)
        style.configure(".", font=default_font, background=BG)

        style.configure(
            "Accent.TButton",
            foreground="white",
            background=ACCENT,
            padding=8,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
        )
        style.map("Accent.TButton", background=[("active", ACCENT_DARK)])

        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_DARK), ("disabled", "#A0C4D6")],
        )

        style.configure(
            "Treeview",
            background=CARD,
            fieldbackground=CARD,
            foreground=TEXT,
            rowheight=28,
        )

        style.configure("Treeview.Heading", background=ACCENT, foreground="white")

        style.configure(
            "TButton",
            foreground=TEXT,  # ✅ IMPORTANT
            background=CARD,
            padding=6,
        )

        style.map(
            "TButton",
            foreground=[("active", "white")],
            background=[("active", ACCENT)],
        )

        style.map("Treeview.Heading", background=[("active", ACCENT_DARK)])

    def _load_or_create_df(self, filename):

        if os.path.exists(filename):
            try:
                self.df = pd.read_excel(filename, dtype=str)
            except Exception:
                self.df = pd.DataFrame()
            required = [
                "RecordID",
                "Block",
                "Name",
                "Date",
                "Year",
                "Month",
                "Room",
                "Wtr Units",
                "Wtr Rate",
                "Wtr Chg",
                "Elec Units",
                "Elec Rate",
                "Elec Chg",
                "Net Units",
                "Net Rate",
                "Net Chg",
                "Misc",
                "Balance",
                "Rent",
                "Total",
                "Paid",
                "Outstanding",
                "Notes",
                "Advance",
                "Credit",
            ]

            for col in required:
                if col not in self.df.columns:
                    self.df[col] = (
                        ""
                        if col
                        in [
                            "RecordID",
                            "Block",
                            "Name",
                            "Date",
                            "Year",
                            "Month",
                            "Room",
                        ]
                        else 0
                    )

            if self.df["Year"].isnull().any() or (self.df["Year"] == "").any():
                self.df["Year"] = self.df["Year"].fillna(str(current_nepali_year()))

            for col in [
                "Wtr Units",
                "Wtr Rate",
                "Wtr Chg",
                "Elec Units",
                "Elec Rate",
                "Elec Chg",
                "Net Units",
                "Net Rate",
                "Net Chg",
                "Misc",
                "Balance",
                "Rent",
                "Total",
                "Paid",
                "Outstanding",
                "Advance",
            ]:
                self.df[col] = (
                    pd.to_numeric(self.df[col], errors="coerce").fillna(0).astype(int)
                )
            if self.df["Year"].isnull().any() or (self.df["Year"] == "").any():

                def infer_year(r):
                    return str(current_nepali_year())

                self.df["Year"] = self.df.apply(infer_year, axis=1)
        else:
            self.df = pd.DataFrame(
                columns=[
                    "RecordID",
                    "Block",
                    "Name",
                    "Date",
                    "Year",
                    "Month",
                    "Room",
                    "Wtr Units",
                    "Wtr Rate",
                    "Wtr Chg",
                    "Elec Units",
                    "Elec Rate",
                    "Elec Chg",
                    "Net Units",
                    "Net Rate",
                    "Net Chg",
                    "Misc",
                    "Balance",
                    "Rent",
                    "Total",
                    "Paid",
                    "Outstanding",
                    "Notes",
                    "Advance",
                ]
            )
            self._save_df()

        if "Credit" in self.df.columns:
            self.df["Credit"] = self.df["Credit"].astype(str)
            self.df["Credit"] = self.df["Credit"].replace("nan", "")

        self.df["RecordID"] = self.df["RecordID"].astype(str)
        self.df["Year"] = self.df["Year"].astype(str)
        self.df["Month"] = self.df["Month"].astype(str)
        self.df["Advance"] = (
            pd.to_numeric(self.df["Advance"], errors="coerce").fillna(0).astype(int)
        )

        self.df["RecordID"] = self.df["RecordID"].astype(str)

        if self.df["RecordID"].duplicated().any():
            print("⚠ Duplicate IDs found → fixing...")
            self.df["RecordID"] = [str(uuid.uuid4()) for _ in range(len(self.df))]

    def _year_range(self, start=None, count=20):
        if start is None:
            start = current_nepali_year() - 10  # show nearby years
        return [str(start + i) for i in range(count)]

    def _save_df(self):
        self.set_sync_status("⏳ Saving...", "#FACC15")

        try:
            if "RecordID" not in self.df.columns:
                self.df["RecordID"] = [str(uuid.uuid4()) for _ in range(len(self.df))]

            # ✅ BACKUP INTO SINGLE FILE
            backup_file = "backups.xlsx"
            sheet_name = datetime.now().strftime("backup_%Y%m%d_%H%M%S")

            try:

                if os.path.exists(backup_file):
                    with pd.ExcelWriter(
                        backup_file,
                        engine="openpyxl",
                        mode="a",
                        if_sheet_exists="overlay",
                    ) as writer:

                        # ✅ Keep max 20 sheets
                        book = writer.book
                        sheets = book.sheetnames

                        if len(sheets) > 20:
                            for s in sheets[:-20]:
                                try:
                                    del writer.book[s]
                                except:
                                    pass

                        self.df.to_excel(writer, sheet_name=sheet_name, index=False)

                else:
                    # create new backup file
                    with pd.ExcelWriter(backup_file, engine="openpyxl") as writer:
                        self.df.to_excel(writer, sheet_name=sheet_name, index=False)

            except Exception as e:
                print("Backup error:", e)

            # ✅ Ensure all RecordIDs are unique
            self.df["RecordID"] = self.df["RecordID"].astype(str)

            self.df = self.df.drop_duplicates(subset=["RecordID"], keep="last")

            # ✅ Save main data
            if "Credit" in self.df.columns:
                self.df["Credit"] = self.df["Credit"].astype(str)

            temp_file = self.filename.replace(".xlsx", "_temp.xlsx")
            self.df.to_excel(temp_file, index=False)
            os.replace(temp_file, self.filename)

            self._build_index()
            self._build_month_cache()

            self.set_sync_status("✅ Synced", "#22C55E")

        except PermissionError:
            self.set_sync_status("❌ Error", "#EF4444")
            messagebox.showerror(
                "Excel File Open",
                "Please close tenants.xlsx and try again.",
            )

    def add_tenant(self):
        self.set_sync_status("⏳ Adding...", "#FACC15")
        try:
            name = self.entries["Name"].get().strip()
            if not name:
                self.set_sync_status("❌ Error", "#EF4444")
                messagebox.showerror("Error", "Tenant name is required")
                return
            block = self.entries["Block"].get().strip()
            year = self.entries["Year"].get().strip() or str(current_nepali_year())
            month = self.entries["Month"].get().strip()
            room = self.entries["Room"].get().strip()
            wtr_units = self.safe_int(self.entries["Wtr Units"].get() or 0)
            wtr_rate = self.safe_int(self.entries["Wtr Rate"].get() or 0)
            elec_units = self.safe_int(self.entries["Elec Units"].get() or 0)
            elec_rate = self.safe_int(self.entries["Elec Rate"].get() or 0)
            net_units = self.safe_int(self.entries["Net Units"].get() or 0)
            net_rate = self.safe_int(self.entries["Net Rate"].get() or 0)
            misc = self.safe_int(self.entries["Misc"].get() or 0)
            balance = self.safe_int(self.entries["Balance"].get() or 0)
            paid = self.safe_int(self.entries["Paid"].get() or 0)
            notes = self.entries["Notes"].get("1.0", tk.END).strip()
            credits = self.entries["Credit"].get().strip()
            self.entries["Notes"].delete("1.0", tk.END)

            # ✅ DUPLICATE CHECK (add here)

            exists = self.df[
                (self.df["Name"] == name)
                & (self.df["Room"] == room)
                & (self.df["Year"] == str(year))
                & (self.df["Month"] == month)
            ]

            if not exists.empty:
                messagebox.showwarning(
                    "Duplicate", "Tenant already exists for this month"
                )
                self.set_sync_status("✅ Synced", "#22C55E")
                return

            date = current_nepali_date_str()
            wtr_chg = wtr_units * wtr_rate
            elec_chg = elec_units * elec_rate
            net_chg = net_units * net_rate
            rent = self.monthly_rent
            total = rent + wtr_chg + net_chg + misc + balance + elec_chg

            advance = 0  # initial base advance
            outstanding = 0
            record_id = str(uuid.uuid4())
            new_record = {
                "RecordID": record_id,
                "Block": block,
                "Name": name,
                "Date": date,
                "Year": str(year),
                "Month": month,
                "Room": room,
                "Wtr Units": wtr_units,
                "Wtr Rate": wtr_rate,
                "Wtr Chg": wtr_chg,
                "Elec Units": elec_units,
                "Elec Rate": elec_rate,
                "Elec Chg": elec_chg,
                "Net Units": net_units,
                "Net Rate": net_rate,
                "Net Chg": net_chg,
                "Misc": misc,
                "Balance": balance,
                "Rent": rent,
                "Total": total,
                "Paid": paid,
                "Outstanding": outstanding,
                "Advance": advance,
                "Notes": notes,
                "Credit": credits,
            }

            for col in [
                "Wtr Units",
                "Wtr Rate",
                "Wtr Chg",
                "Elec Units",
                "Elec Rate",
                "Elec Chg",
                "Net Units",
                "Net Rate",
                "Net Chg",
                "Misc",
                "Balance",
                "Rent",
                "Total",
                "Paid",
                "Outstanding",
                "Advance",
            ]:
                self.df[col] = (
                    pd.to_numeric(self.df[col], errors="coerce").fillna(0).astype(int)
                )
            self._pending_navigation = (str(year), str(month))

            self.df = pd.concat(
                [self.df, pd.DataFrame([new_record])], ignore_index=True
            )

            idx = self.df.index[-1]
            rid = str(self.df.at[idx, "RecordID"])

            # ✅ IMPORTANT: register base advance
            self._base_advance_map[rid] = 0

            # ✅ Recalculate using system
            total, outstanding, new_advance = self.calculate_totals(idx)

            self.df.at[idx, "Total"] = total
            self.df.at[idx, "Outstanding"] = outstanding
            self.df.at[idx, "Advance"] = new_advance
            self._save_df()
            self.set_sync_status("✅ Synced", "#22C55E")
            messagebox.showinfo("Success", f"Tenant {name} added successfully!")

            for entry in self.entries.values():
                try:
                    entry.delete(0, tk.END)
                except Exception:
                    pass
            self.entries["Year"].set(str(current_nepali_year()))
            self.entries["Month"].set(NEPALI_MONTHS[0])
            self.entries["Paid"].insert(0, "0")

            if self.show_all_window and tk.Toplevel.winfo_exists(self.show_all_window):
                self.refresh_show_all()

        except ValueError:
            messagebox.showerror(
                "Error",
                "Please enter valid numeric values for units/rates/misc/balance/paid.",
            )

    def delete_tenant(self, row):
        self.set_sync_status("⏳ Updating...", "#FACC15")
        confirm = messagebox.askyesno("Delete", f"Delete {row['Name']}?")
        if not confirm:
            self.set_sync_status("✅ Synced", "#22C55E")
            return

        rid = str(row["RecordID"])

        # ✅ remove from dataframe
        self.df = self.df[self.df["RecordID"] != rid]
        self._save_df()

        # ✅ remove from UI instantly
        year = str(row["Year"])
        month = str(row["Month"])

        remaining = self.df[
            (self.df["Year"].astype(str) == str(year))
            & (self.df["Month"].astype(str) == str(month))
        ]

        ui = self._ui_map.get(rid)
        if ui:
            try:
                ui["frame"].destroy()
            except:
                pass
            self._ui_map.pop(rid, None)
            self._last_data_hash.pop(rid, None)

            # ✅ update summary immediately

        try:
            total_paid = int(remaining["Paid"].sum())
            summary_map = self._ui_map.get("_summary", {})
            lbl = summary_map.get((year, month))
            if lbl:
                lbl.config(text=f"Total Collected: {total_paid}")
        except:
            pass

        if remaining.empty:

            self.remove_month_from_sidebar(year, month)

            key = (year, month)
            frame = self._frame_map.get(key)

            if frame:
                try:
                    frame.destroy()
                except:
                    pass
                self._frame_map.pop(key, None)

            # ✅ ONLY skip popup, not logic
            if not self._is_rebuilding_ui:
                messagebox.showinfo(
                    "Deleted",
                    f"{month} {year} has no tenants left ✅",
                )

            self.set_sync_status("✅ Synced", "#22C55E")
            return

    def show_all(self):

        if self.show_all_window and tk.Toplevel.winfo_exists(self.show_all_window):
            self.refresh_show_all()

            return

        top = tk.Toplevel(self.root)
        top.protocol("WM_DELETE_WINDOW", self.root.destroy)
        top.title("All Tenants")
        top.state("zoomed")
        self.show_all_window = top
        top.configure(bg=BG)

        self._ui_map.clear()
        self._frame_map.clear()
        self._last_selected_frame_key = None

        paned = ttk.Panedwindow(top, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True)

        sidebar_frame = tk.Frame(paned, bg=BG, width=320)
        paned.add(sidebar_frame, weight=0)

        # Scrollable content area
        content_scroll_frame = tk.Frame(paned, bg=BG)
        paned.add(content_scroll_frame, weight=1)

        canvas = tk.Canvas(content_scroll_frame, bg=BG, highlightthickness=0)
        self.canvas = canvas

        canvas.configure(yscrollincrement=20)

        def limit_scroll(*args):
            canvas.yview(*args)

            # ✅ clamp between 0 and 1
            start, end = canvas.yview()

            if start <= 0:
                canvas.yview_moveto(0)
            elif end >= 1:
                canvas.yview_moveto(1 - (end - start))

        scrollbar = ttk.Scrollbar(
            content_scroll_frame, orient="vertical", command=limit_scroll
        )

        content_frame = tk.Frame(canvas, bg=BG)
        # ✅ GLOBAL TOP SEARCH BAR (ONLY ONCE)
        topbar = tk.Frame(content_frame, bg=BG)
        topbar.grid(row=0, column=0, sticky="ew", pady=(8, 6))

        ttk.Button(topbar, text="⚙ Settings", command=self.open_settings).pack(
            side="right", padx=6
        )

        ttk.Button(
            topbar,
            text="Detailed Print Selected",
            command=self.detailed_print_selected,
            style="Accent.TButton",
        ).pack(side="left", padx=6)

        ttk.Button(
            topbar,
            text="Basic Print Selected",
            command=self.print_selected,
            style="Accent.TButton",
        ).pack(side="left", padx=6)

        ttk.Button(
            topbar,
            text="Select All",
            command=self.toggle_select_all,
            style="Accent.TButton",
        ).pack(side="left", padx=6)

        ttk.Button(
            topbar,
            text="Add New Tenant",
            command=self.add_new_dialog,
            style="Accent.TButton",
        ).pack(side="left", padx=6)

        # push search to right
        spacer = tk.Frame(topbar, bg=BG)
        spacer.pack(side="left", expand=True)

        tk.Label(topbar, text="🔍", bg=BG, fg=TEXT).pack(side="left", padx=(10, 2))

        self.global_search_entry = ttk.Entry(
            topbar, textvariable=self.tenant_search_var, width=30
        )
        self.global_search_entry.pack(side="left", padx=(0, 10))

        self.global_search_entry.insert(0, "Search tenant...")
        self.global_search_entry.config(foreground="gray")

        def clear_placeholder(event):
            if self.global_search_entry.get() == "Search tenant...":
                self.global_search_entry.delete(0, tk.END)
                self.global_search_entry.config(foreground="white")

        def restore_placeholder(event):
            if not self.global_search_entry.get():
                self.global_search_entry.insert(0, "Search tenant...")
                self.global_search_entry.config(foreground="gray")

        self.global_search_entry.bind("<FocusIn>", clear_placeholder)
        self.global_search_entry.bind("<FocusOut>", restore_placeholder)

        def update_scroll_region(event=None):
            canvas.update_idletasks()
            bbox = canvas.bbox("all")
            if bbox:
                canvas.configure(scrollregion=(0, 0, bbox[2], bbox[3]))

        content_frame.bind("<Configure>", update_scroll_region)

        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.sync_status = tk.Label(
            topbar,
            text=f"✅ Synced • {self.current_time_str()}",
            bg=BG,
            fg="#22C55E",
            font=("Segoe UI", 9, "bold"),
        )
        self.sync_status.pack(side="right", padx=10)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            # ✅ clamp after scroll
            start, end = canvas.yview()
            if start <= 0:
                canvas.yview_moveto(0)
            elif end >= 1:
                canvas.yview_moveto(1 - (end - start))

        canvas.bind(
            "<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel)
        )
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        ttk.Label(
            sidebar_frame, text="Navigate Year → Month", font=("Segoe UI", 11, "bold")
        ).pack(anchor="w", padx=12, pady=(12, 6))

        # Date search (Year/Month)
        ttk.Label(sidebar_frame, text="Search Year/Month").pack(anchor="w", padx=12)
        date_entry = ttk.Entry(
            sidebar_frame, textvariable=self.date_search_var, width=28
        )
        date_entry.pack(fill="x", padx=12, pady=(0, 8))

        # Add New Tenant button in sidebar (prominent)
        add_new_btn = ttk.Button(
            sidebar_frame,
            text="Add New Tenant",
            command=self.add_new_dialog,
            style="Accent.TButton",
        )
        add_new_btn.pack(fill="x", padx=12, pady=(0, 8))

        # Keep the sidebar Batch Carry Over button (user requested)
        batch_btn = ttk.Button(
            sidebar_frame,
            text="Batch Carry Over",
            command=lambda: self.batch_carry_over_from_sidebar(sidebar_tree),
        )
        batch_btn.pack(fill="x", padx=12, pady=(0, 8))

        sidebar_tree = ttk.Treeview(sidebar_frame, show="tree", height=24)
        sidebar_tree.pack(fill="both", expand=True, padx=12, pady=6)
        self.sidebar_tree = sidebar_tree
        # Build combos
        self.df["Year"] = self.df["Year"].astype(str)

        # ✅ only months with actual tenants
        valid_df = self.df[self.df["Name"].astype(str).str.strip() != ""]

        combos = (
            valid_df[["Year", "Month"]].drop_duplicates().fillna("").values.tolist()
        )

        def combo_key(c):
            y, m = c
            try:
                mi = NEPALI_MONTHS.index(m) if m in NEPALI_MONTHS else -1
            except Exception:
                mi = -1
            try:
                yi = int(y)
            except Exception:
                yi = current_nepali_year()

            return (yi, mi)

        combos_sorted = sorted(combos, key=combo_key)

        year_months = {}

        # ✅ build mapping
        for y, m in combos_sorted:
            y_str = str(y)
            m_str = str(m)
            year_months.setdefault(y_str, []).append(m_str)

        # ✅ build sidebar
        for year, months in year_months.items():
            year_node = sidebar_tree.insert(
                "", "end", iid=f"year_{year}", text=year, open=False
            )

            for month in months:
                child_iid = f"ym_{year}_{month or 'Unspecified'}"

                sidebar_tree.insert(year_node, "end", iid=child_iid, text=month)

                self._frame_map[(str(year), str(month))] = None

        content_frame.grid_columnconfigure(0, weight=1)

        def on_sidebar_select(event):
            sel = sidebar_tree.focus()

            if not sel:
                return

            if sel.startswith("ym_"):
                parts = sel.split("_", 2)
                if len(parts) < 3:
                    return

                year = parts[1]
                month = parts[2]

                month_val = "" if month == "Unspecified" else month
                key = (str(year), str(month_val))

                frame = self._frame_map.get(key)

                # ✅ CREATE FRAME ONLY WHEN CLICKED
                if frame is None:
                    frame = tk.Frame(content_frame, bg=BG)
                    frame.grid(row=1, column=0, sticky="nsew")

                    self._frame_map[key] = frame

                    self.build_month_ui(frame, year, month_val)

                # ✅ hide all frames
                for f in self._frame_map.values():
                    if f:
                        f.grid_remove()

                frame.grid()
                self._last_selected_frame_key = key

                if hasattr(self, "canvas"):
                    self.canvas.update_idletasks()
                    self.canvas.yview_moveto(0)

            else:
                sidebar_tree.item(sel, open=not sidebar_tree.item(sel, "open"))

        sidebar_tree.bind("<<TreeviewSelect>>", on_sidebar_select)
        # ✅ AUTO-OPEN LAST ADDED MONTH

        if self._pending_navigation:

            def go_to_added():
                try:
                    y, m = self._pending_navigation
                    node_id = f"ym_{y}_{m}"

                    if not sidebar_tree.exists(node_id):
                        self.root.after_idle(app.show_all)
                        return

                    parent_id = f"year_{y}"

                    sidebar_tree.item(parent_id, open=True)

                    sidebar_tree.selection_set(node_id)
                    sidebar_tree.focus(node_id)

                    # call directly
                    sidebar_tree.event_generate("<<TreeviewSelect>>")

                    # ✅ CLEAR AFTER SUCCESS
                    self._pending_navigation = None

                except Exception:
                    pass

            self.root.after(400, go_to_added)

    def refresh_show_all(self):
        self._is_rebuilding_ui = True

        self._load_or_create_df(self.filename)
        if self.show_all_window and tk.Toplevel.winfo_exists(self.show_all_window):
            try:
                self.show_all_window.destroy()
            except:
                pass

        self.root.after(50, self.show_all)

        self._is_rebuilding_ui = False

    def refresh_current_frame(self):

        if not self._ui_map:
            return
        self.root.update_idletasks()

        for rid, ui in self._ui_map.items():
            if rid == "_summary":
                continue

            idx = self._id_index.get(rid)
            if idx is None:
                continue

            row = self.df.iloc[idx]

            notes_val = str(row["Notes"]).strip()
            ui["notes_label"].config(
                text=f"📝 {notes_val or '(No notes)'}",
                fg="#CBD5F5" if notes_val else "#64748B",
            )

            # ✅ Create hash of important fields

            new_hash = (
                self.safe_int(row["Wtr Units"]),
                self.safe_int(row["Wtr Rate"]),
                self.safe_int(row["Elec Units"]),
                self.safe_int(row["Elec Rate"]),
                self.safe_int(row["Net Units"]),
                self.safe_int(row["Net Rate"]),
                self.safe_int(row["Misc"]),
                self.safe_int(row["Balance"]),
                self.safe_int(row["Advance"]),
                self.safe_int(row["Total"]),
                self.safe_int(row["Paid"]),
                self.safe_int(row["Outstanding"]),
            )

            # ✅ Skip if nothing changed
            if self._last_data_hash.get(rid) == new_hash:
                continue

            # ✅ Save new hash
            self._last_data_hash[rid] = new_hash

            tree = ui["tree"]

            try:
                # ✅ temporarily disable redraw (hacky but effective)

                tree.configure(takefocus=False)
                tree.after_idle(lambda: tree.update_idletasks())
                # ✅ batch update (no repaint in between)

                tree.item(
                    rid + "_elec",
                    values=(
                        "Electricity",
                        self.safe_int(row["Elec Units"]),
                        self.safe_int(row["Elec Rate"]),
                        self.safe_int(row["Elec Chg"]),
                    ),
                )

                tree.item(
                    rid + "_net",
                    values=(
                        "Internet",
                        self.safe_int(row["Net Units"]),
                        self.safe_int(row["Net Rate"]),
                        self.safe_int(row["Net Chg"]),
                    ),
                )

                tree.item(
                    rid + "_water",
                    values=(
                        "Water",
                        self.safe_int(row["Wtr Units"]),
                        self.safe_int(row["Wtr Rate"]),
                        self.safe_int(row["Wtr Chg"]),
                    ),
                )

                tree.item(
                    rid + "_misc", values=("Misc", "-", "-", self.safe_int(row["Misc"]))
                )
                tree.item(
                    rid + "_bal",
                    values=("Old Balance", "-", "-", self.safe_int(row["Balance"])),
                )
                tree.item(
                    rid + "_advance",
                    values=("Advance (Auto)", "-", "-", self.safe_int(row["Advance"])),
                )

                tree.item(
                    rid + "_total",
                    values=("Total Bill", "-", "-", self.safe_int(row["Total"])),
                )

                credit_val = self.clean_credit(row.get("Credit", ""))

                tree.item(
                    rid + "_paid",
                    values=("Paid", credit_val or "-", "-", self.safe_int(row["Paid"])),
                )

                tree.item(
                    rid + "_outstanding",
                    values=(
                        "Outstanding",
                        "-",
                        "-",
                        self.safe_int(row["Outstanding"]),
                    ),
                )

                # ✅ restore redraw
                tree.configure(takefocus=True)
                tree.update_idletasks()

                # ✅ update label AFTER redraw restore
                ui["total_label"].config(
                    text=f"Tenant Total: {self.safe_int(row['Total'])}"
                )

                # ✅ UPDATE COLOR LIVE
                frame = ui["frame"]

                self.set_frame_color(frame, self.get_row_color(row))

            except Exception:
                pass

        self._build_month_cache()
        # ✅ Update summaries after refresh
        try:
            summary_map = self._ui_map.get("_summary", {})

            for (year, month), lbl in summary_map.items():
                month_rows = self._month_cache.get((year, month), [])
                total_paid = sum(self.safe_int(r.get("Paid", 0)) for r in month_rows)

                lbl.config(text=f"Total Collected: {total_paid}")

        except Exception:
            pass

    # safely rebuild

    def add_new_dialog(self):
        """
        Modal dialog to add a new tenant quickly from the dashboard.
        """
        dlg = tk.Toplevel(self.show_all_window or self.root)
        dlg.title("Add New Tenant")

        dlg.geometry("380x580")  # ✅ bigger
        dlg.configure(bg="#0F172A")
        # ✅ Scroll setup
        canvas = tk.Canvas(dlg, bg="#0F172A", highlightthickness=0)
        scrollbar = ttk.Scrollbar(dlg, orient="vertical", command=canvas.yview)

        scroll_frame = tk.Frame(canvas, bg="#0F172A")

        scroll_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind(
            "<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_mousewheel)
        )
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        dlg.update_idletasks()
        x = (dlg.winfo_screenwidth() // 2) - (480 // 2)
        y = (dlg.winfo_screenheight() // 2) - (620 // 2)
        dlg.geometry(f"+{x}+{y}")  # ✅ center

        dlg.transient(self.show_all_window or self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        widgets = {}

        # ✅ Section builder
        def section(title):
            card = tk.Frame(scroll_frame, bg="#1E293B", bd=1, relief="solid")
            card.pack(fill="x", padx=40, pady=8)

            header = tk.Label(
                card,
                text=title,
                bg="#1E293B",
                fg="#60A5FA",
                font=("Segoe UI", 11, "bold"),
            )
            header.pack(anchor="w", padx=10, pady=(8, 4))

            body = tk.Frame(card, bg="#1E293B")
            body.pack(fill="x", padx=10, pady=(0, 10))

            return body

        # ✅ Field builder
        def add_field(parent, label, field):
            rowf = tk.Frame(parent, bg="#1E293B")
            rowf.pack(fill="x", pady=4)

            tk.Label(
                rowf,
                text=label,
                bg="#1E293B",
                fg="#E2E8F0",
                width=16,
                anchor="w",
            ).pack(side="left")

            # ✅ Get current selected frame (year, month)
            default_year = str(current_nepali_year())
            default_month = NEPALI_MONTHS[0]

            if (
                hasattr(self, "_last_selected_frame_key")
                and self._last_selected_frame_key
            ):
                y, m = self._last_selected_frame_key
                if y:
                    default_year = str(y)
                if m:
                    default_month = str(m)

            if field == "Year":
                e = ttk.Combobox(rowf, values=self._year_range(), width=20)
                e.set(default_year)

            elif field == "Month":
                e = ttk.Combobox(rowf, values=NEPALI_MONTHS, width=20)
                e.set(default_month)
            else:
                e = ttk.Entry(rowf, width=22)

            e.pack(side="left", fill="x", expand=True)
            widgets[field] = e

        # ✅ Build sections

        # 👤 Basic Info
        basic = section("👤 Basic Info")
        add_field(basic, "Block", "Block")
        add_field(basic, "Name", "Name")
        add_field(basic, "Room", "Room")

        # 📅 Date
        date = section("📅 Date")
        add_field(date, "Year", "Year")
        add_field(date, "Month", "Month")

        # 💡 Utilities
        util = section("💡 Utilities")
        add_field(util, "Water Units", "Wtr Units")
        add_field(util, "Water Rate", "Wtr Rate")
        add_field(util, "Electric Units", "Elec Units")
        add_field(util, "Electric Rate", "Elec Rate")
        add_field(util, "Net Units", "Net Units")
        add_field(util, "Net Rate", "Net Rate")

        # 💰 Financial
        fin = section("💰 Financial")
        add_field(fin, "Misc", "Misc")
        add_field(fin, "Balance", "Balance")
        add_field(fin, "Paid", "Paid")
        add_field(fin, "Credit", "Credit")

        widgets["Name"].focus()
        widgets["Name"].select_range(0, tk.END)

        def on_ok():
            try:
                block = widgets["Block"].get().strip()
                name = widgets["Name"].get().strip()
                year = widgets["Year"].get().strip() or str(current_nepali_year())
                month = widgets["Month"].get().strip()
                room = widgets["Room"].get().strip()

                wtr_units = self.safe_int(widgets["Wtr Units"].get() or 0)
                wtr_rate = self.safe_int(widgets["Wtr Rate"].get() or 0)
                elec_units = self.safe_int(widgets["Elec Units"].get() or 0)
                elec_rate = self.safe_int(widgets["Elec Rate"].get() or 0)
                net_units = self.safe_int(widgets["Net Units"].get() or 0)
                net_rate = self.safe_int(widgets["Net Rate"].get() or 0)

                misc = self.safe_int(widgets["Misc"].get() or 0)
                balance = self.safe_int(widgets["Balance"].get() or 0)
                paid = self.safe_int(widgets["Paid"].get() or 0)
                credit = widgets["Credit"].get().strip()

                exists = self.df[
                    (self.df["Name"] == name)
                    & (self.df["Room"] == room)
                    & (self.df["Year"] == str(year))
                    & (self.df["Month"] == month)
                ]

                if not exists.empty:
                    messagebox.showwarning(
                        "Duplicate",
                        f"{name} already exists in Room {room} for this month",
                    )
                    return

                date = current_nepali_date_str()

                wtr_chg = wtr_units * wtr_rate
                net_chg = net_units * net_rate
                elec_chg = elec_units * elec_rate
                rent = self.monthly_rent

                total = rent + wtr_chg + net_chg + elec_chg + misc + balance

                advance = 0
                outstanding = 0

                record_id = str(uuid.uuid4())

                new_record = {
                    "RecordID": record_id,
                    "Block": block,
                    "Name": name,
                    "Date": date,
                    "Year": str(year),
                    "Month": month,
                    "Room": room,
                    "Wtr Units": wtr_units,
                    "Wtr Rate": wtr_rate,
                    "Wtr Chg": wtr_chg,
                    "Net Units": net_units,
                    "Net Rate": net_rate,
                    "Net Chg": net_chg,
                    "Elec Units": elec_units,
                    "Elec Rate": elec_rate,
                    "Elec Chg": elec_chg,
                    "Misc": misc,
                    "Balance": balance,
                    "Rent": rent,
                    "Total": total,
                    "Paid": paid,
                    "Outstanding": outstanding,
                    "Advance": advance,
                    "Credit": credit,
                }

                self.df = pd.concat(
                    [self.df, pd.DataFrame([new_record])], ignore_index=True
                )

                for col in [
                    "Wtr Units",
                    "Wtr Rate",
                    "Wtr Chg",
                    "Net Units",
                    "Net Rate",
                    "Net Chg",
                    "Elec Units",
                    "Elec Rate",
                    "Elec Chg",
                    "Misc",
                    "Balance",
                    "Rent",
                    "Total",
                    "Paid",
                    "Outstanding",
                    "Advance",
                ]:
                    self.df[col] = (
                        pd.to_numeric(self.df[col], errors="coerce")
                        .fillna(0)
                        .astype(int)
                    )
                self._build_index()

                idx = self.df.index[-1]
                rid = str(self.df.at[idx, "RecordID"])

                self._base_advance_map[rid] = 0

                total, outstanding, new_advance = self.calculate_totals(idx)

                self.df.at[idx, "Total"] = total
                self.df.at[idx, "Outstanding"] = outstanding
                self.df.at[idx, "Advance"] = new_advance
                self._save_df()

                dlg.destroy()
                messagebox.showinfo("Success", f"Tenant {name} added.")

                if self.show_all_window and tk.Toplevel.winfo_exists(
                    self.show_all_window
                ):
                    self.refresh_show_all()

            except ValueError:
                messagebox.showerror(
                    "Error",
                    "Please enter valid numeric values.",
                    parent=dlg,
                )

        def on_cancel():
            dlg.destroy()

        btn_frame = tk.Frame(scroll_frame, bg="#0F172A")
        btn_frame.pack(pady=(6, 8))
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(
            side="left", padx=6
        )
        dlg.bind("<Return>", lambda e: on_ok())
        dlg.wait_window()

    def carry_over_tenant_dialog(self, row):
        cur_month = str(row["Month"]).strip()

        try:
            cur_year = int(row["Year"])
        except Exception:
            cur_year = current_nepali_year()

        # ✅ Next month logic
        if cur_month in NEPALI_MONTHS:
            idx = NEPALI_MONTHS.index(cur_month)
            next_idx = (idx + 1) % 12
            default_month = NEPALI_MONTHS[next_idx]
            default_year = cur_year + 1 if next_idx == 0 else cur_year
        else:
            default_month = NEPALI_MONTHS[0]
            default_year = cur_year

        # ✅ Dialog UI
        dlg = tk.Toplevel(self.show_all_window or self.root)
        dlg.title("Carry Over Tenant")
        dlg.transient(self.show_all_window or self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        tk.Label(dlg, text="Select Year:").grid(
            row=0, column=0, padx=8, pady=8, sticky="w"
        )
        year_cb = ttk.Combobox(
            dlg,
            values=self._year_range(),
            state="normal",
            width=18,
        )
        year_cb.grid(row=0, column=1, padx=8, pady=8)
        year_cb.set(str(default_year))

        tk.Label(dlg, text="Select Nepali Month:").grid(
            row=1, column=0, padx=8, pady=8, sticky="w"
        )
        month_cb = ttk.Combobox(dlg, values=NEPALI_MONTHS, state="normal", width=18)
        month_cb.grid(row=1, column=1, padx=8, pady=8)
        month_cb.set(default_month)

        result = {"ok": False, "year": None, "month": None}

        def on_ok():
            sel_year = year_cb.get().strip()
            sel_month = month_cb.get().strip()

            if not sel_year or not sel_month:
                messagebox.showwarning(
                    "Missing", "Please select both year and month.", parent=dlg
                )
                return

            result["ok"] = True
            result["year"] = sel_year
            result["month"] = sel_month
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg="#0F172A")
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(0, 8))
        ttk.Button(btn_frame, text="OK", command=on_ok).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).grid(
            row=0, column=1, padx=6
        )

        dlg.wait_window()

        if not result["ok"]:
            return

        # ✅ CREATE NEW ENTRY (AFTER dialog)
        try:
            self._save_df()

            rec_id = str(row["RecordID"])

            fresh_rows = self.df[self.df["RecordID"].astype(str) == rec_id]

            if fresh_rows.empty:
                messagebox.showerror("Error", "Source record not found.")
                return

            fresh = fresh_rows.iloc[-1]

            exists = self.df[
                (self.df["Name"] == fresh["Name"])
                & (self.df["Room"] == fresh["Room"])
                & (self.df["Year"] == str(result["year"]))
                & (self.df["Month"] == result["month"])
            ]

            if not exists.empty:
                messagebox.showwarning(
                    "Duplicate",
                    f"{fresh['Name']} already exists in Room {fresh['Room']} for this month",
                )
                return

            prev_outstanding = self.safe_int(fresh["Outstanding"])
            prev_advance = self.safe_int(fresh.get("Advance", 0))

            # ✅ NET POSITION
            net = prev_outstanding - prev_advance

            if net > 0:
                carry_balance = net
                new_advance = 0
            else:
                carry_balance = 0
                new_advance = abs(net)

            new_advance = max(new_advance, 0)
            carry_balance = max(carry_balance, 0)

            # ✅ Smart carry for Water

            # ✅ Water
            wtr_rate = self.safe_int(fresh.get("Wtr Rate", 0))

            # ✅ Electricity (FIX)
            elec_rate = self.safe_int(fresh.get("Elec Rate", 0))

            # ✅ Internet (FIX with optional logic)
            if self.safe_int(fresh.get("Net Units", 0)) > 0:
                net_rate = self.safe_int(fresh.get("Net Rate", 0))
            else:
                net_rate = 0

            new_record = {
                "RecordID": str(uuid.uuid4()),
                "Block": fresh["Block"],
                "Name": fresh["Name"],
                "Date": current_nepali_date_str(),
                "Year": str(result["year"]),
                "Month": result["month"],
                "Room": fresh["Room"],
                # ✅ RESET UNITS (NEW MONTH USAGE)
                "Wtr Units": 0,
                "Wtr Rate": wtr_rate,
                "Wtr Chg": 0,
                "Elec Units": 0,
                "Elec Rate": elec_rate,
                "Elec Chg": 0,
                "Net Units": 0,
                "Net Rate": net_rate,
                "Net Chg": 0,
                # ✅ RESET EXTRA CHARGES
                "Misc": 0,
                # ✅ KEEP PREVIOUS BALANCE
                "Balance": carry_balance,
                "Advance": new_advance,
                "Rent": self.monthly_rent,
                # ✅ RESET PAYMENT
                "Paid": 0,
                "Notes": fresh.get("Notes", ""),
                "Credit": fresh.get("Credit", ""),
            }

            # ✅ insert once
            self.df = pd.concat(
                [self.df, pd.DataFrame([new_record])], ignore_index=True
            )

            idx = self.df.index[-1]

            rid = str(self.df.at[idx, "RecordID"])

            self._base_advance_map[rid] = self.safe_int(self.df.at[idx, "Advance"])

            base_advance = self.get_base_advance(rid, idx)
            self.df.at[idx, "Advance"] = base_advance

            total, outstanding, new_advance = self.calculate_totals(idx)

            self.df.at[idx, "Total"] = total
            self.df.at[idx, "Outstanding"] = outstanding
            self.df.at[idx, "Advance"] = new_advance

            self._save_df()

            messagebox.showinfo(
                "Success",
                f"{fresh['Name']} carried over to {result['month']} {result['year']}",
            )

            self.refresh_show_all()

        except Exception as e:
            messagebox.showerror("Error", f"Carry over failed: {e}")

    def batch_carry_over_from_sidebar(self, sidebar_tree):
        sel = sidebar_tree.focus()
        if not sel or not sel.startswith("ym_"):
            messagebox.showwarning(
                "Select Month",
                "Please select a month in the sidebar to batch carry over.",
                parent=self.show_all_window,
            )
            return
        parts = sel.split("_", 2)
        if len(parts) < 3:
            messagebox.showwarning(
                "Select Month",
                "Please select a valid month node.",
                parent=self.show_all_window,
            )
            return
        year = parts[1]
        month = parts[2]
        if month == "Unspecified":
            month_val = ""
        else:
            month_val = month
        source_rows = self.df[
            (self.df["Year"].astype(str) == str(year))
            & (self.df["Month"].astype(str) == str(month_val))
        ]
        if source_rows.empty:
            messagebox.showinfo(
                "No Tenants",
                "No tenants found for the selected month.",
                parent=self.show_all_window,
            )
            return

        try:
            cur_year = int(year)

        except Exception:
            cur_year = current_nepali_year()

        cur_month = month_val
        if cur_month in NEPALI_MONTHS:
            idx = NEPALI_MONTHS.index(cur_month)
            next_idx = (idx + 1) % 12
            default_month = NEPALI_MONTHS[next_idx]
            default_year = cur_year + 1 if next_idx == 0 else cur_year
        else:
            default_month = NEPALI_MONTHS[0]
            default_year = cur_year

        dlg = tk.Toplevel(self.show_all_window or self.root)
        dlg.title("Batch Carry Over")
        dlg.transient(self.show_all_window or self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        tk.Label(dlg, text=f"Source: {year} - {month_val or 'Unspecified'}").grid(
            row=0, column=0, columnspan=2, padx=8, pady=(8, 4), sticky="w"
        )
        tk.Label(dlg, text="Target Year:").grid(
            row=1, column=0, padx=8, pady=6, sticky="w"
        )
        year_cb = ttk.Combobox(
            dlg,
            values=self._year_range(),
            state="normal",
            width=18,
        )
        year_cb.grid(row=1, column=1, padx=8, pady=6)
        year_cb.set(str(current_nepali_year()))

        tk.Label(dlg, text="Target Nepali Month:").grid(
            row=2, column=0, padx=8, pady=6, sticky="w"
        )
        month_cb = ttk.Combobox(dlg, values=NEPALI_MONTHS, state="normal", width=18)
        month_cb.grid(row=2, column=1, padx=8, pady=6)
        month_cb.set(default_month)

        result = {"ok": False, "year": None, "month": None}

        def on_ok():
            sel_year = year_cb.get().strip()
            sel_month = month_cb.get().strip()
            if not sel_year or not sel_month:
                messagebox.showwarning(
                    "Missing", "Please select both year and month.", parent=dlg
                )
                return
            result["ok"] = True
            result["year"] = sel_year
            result["month"] = sel_month
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg="#0F172A")
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(0, 8))
        ttk.Button(btn_frame, text="OK", command=on_ok).grid(row=0, column=0, padx=6)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).grid(
            row=0, column=1, padx=6
        )

        dlg.wait_window()
        if not result["ok"]:
            return

        try:
            self._save_df()
            target_year = int(result["year"])
            target_month = result["month"]
            date = current_nepali_date_str()
            rent = self.monthly_rent

            new_records = []
            for _, r in source_rows.iterrows():
                rec_id = str(r.get("RecordID"))
                fresh_rows = self.df[self.df["RecordID"].astype(str) == rec_id]
                if fresh_rows.empty:
                    continue
                fresh = fresh_rows.iloc[-1]

                prev_outstanding = self.safe_int(fresh["Outstanding"])
                prev_advance = self.safe_int(fresh.get("Advance", 0))

                net = prev_outstanding - prev_advance

                if net > 0:
                    carry_balance = net
                    new_advance = 0
                else:
                    carry_balance = 0
                    new_advance = abs(net)

                # ✅ Smart carry logic

                wtr_rate = int(fresh.get("Wtr Rate", 0))
                elec_rate = int(fresh.get("Elec Rate", 0))

                if int(fresh.get("Net Units", 0)) > 0:
                    net_rate = int(fresh.get("Net Rate", 0))
                else:
                    net_rate = 0

                new_record = {
                    "RecordID": str(uuid.uuid4()),
                    "Block": fresh["Block"],
                    "Name": fresh["Name"],
                    "Date": date,
                    "Year": str(target_year),
                    "Month": target_month,
                    "Room": fresh["Room"],
                    "Wtr Units": 0,
                    "Wtr Rate": wtr_rate,
                    "Wtr Chg": 0,
                    "Elec Units": 0,
                    "Elec Rate": elec_rate,
                    "Elec Chg": 0,
                    "Net Units": 0,
                    "Net Rate": net_rate,
                    "Net Chg": 0,
                    "Misc": 0,
                    "Balance": carry_balance,
                    "Advance": new_advance,
                    "Rent": rent,
                    "Total": 0,
                    "Paid": 0,
                    "Outstanding": 0,
                    "Notes": fresh.get("Notes", ""),
                    "Credit": fresh.get("Credit", ""),
                }
                new_records.append(new_record)

            if new_records:
                self.df = pd.concat(
                    [self.df, pd.DataFrame(new_records)], ignore_index=True
                )

                start_idx = len(self.df) - len(new_records)

                for i in range(start_idx, len(self.df)):

                    rid = str(self.df.at[i, "RecordID"])

                    self._base_advance_map[rid] = self.safe_int(
                        self.df.at[i, "Advance"]
                    )

                    base_advance = self.get_base_advance(rid, i)
                    self.df.at[i, "Advance"] = base_advance

                    total, outstanding, new_advance = self.calculate_totals(i)

                    self.df.at[i, "Total"] = total
                    self.df.at[i, "Outstanding"] = outstanding
                    self.df.at[i, "Advance"] = new_advance

                for col in [
                    "Wtr Units",
                    "Wtr Rate",
                    "Wtr Chg",
                    "Elec Units",
                    "Elec Rate",
                    "Elec Chg",
                    "Net Units",
                    "Net Rate",
                    "Net Chg",
                    "Misc",
                    "Balance",
                    "Rent",
                    "Total",
                    "Paid",
                    "Outstanding",
                ]:
                    self.df[col] = (
                        pd.to_numeric(self.df[col], errors="coerce")
                        .fillna(0)
                        .astype(int)
                    )
                self._save_df()

            messagebox.showinfo(
                "Batch Carry Over",
                f"Carried over {len(new_records)} tenants to {target_month} {target_year}.",
                parent=self.show_all_window,
            )
            if self.show_all_window and tk.Toplevel.winfo_exists(self.show_all_window):
                self.refresh_show_all()

        except Exception as e:
            messagebox.showerror(
                "Error", f"Batch carry over failed: {e}", parent=self.show_all_window
            )

    def print_selected(self):
        selected_ids = []

        for rid, data in self._ui_map.items():
            if rid == "_summary":
                continue
            if data.get("selected") and data["selected"].get():
                selected_ids.append(rid)

        if not selected_ids:
            messagebox.showwarning("No selection", "Select tenants first")
            return

        all_bills_html = ""

        for record_id in selected_ids:

            idx = self._id_index.get(record_id)
            if idx is None:
                continue

            row = self.df.iloc[idx]

            credit = self.clean_credit(row.get("Credit", ""))
            credit_text = f" ({credit})" if credit else ""

            bill_html = f"""
            <div class="bill">
              <div>{row['Name']} | Rm {row['Room']} | Blk {row['Block']} | {datetime.now().strftime("%Y-%m-%d %I:%M %p")}</div>
                <div class="total">
                    Misc+Other: {self.safe_int(row['Misc']) + self.safe_int(row['Elec Chg']) + self.safe_int(row['Wtr Chg']) + self.safe_int(row['Net Chg'])}
                    | Total: {self.safe_int(row['Total'])}
                </div>
            </div>"""

            all_bills_html += bill_html

        full_html = f"""
                <!doctype html>
                <html>
                <head>
                <meta charset="utf-8">
                <title>All Bills</title>

                <style>   
                    @media print {{
                        button {{
                            display: none;
                        }}

                        body {{
                            margin: 5px;
                            font-size: 11px;
                        }}
                    }}

                    body {{
                        font-family: Arial;
                        margin: 10px;
                        font-size: 12px;
                    }}

                    /* ✅ 2 COLUMN LAYOUT */
                    .container {{
                        display: grid;
                        grid-template-columns: repeat(2, 1fr);
                        gap: 9px;
                    }}

                    /* ✅ EACH BILL */
                    .bill {{
                        line-height: 1.1;
                        font-size: 12px;
                                            
                        margin-bottom: 14px;   /* ✅ adds space between rows */
                        padding-bottom: 6px;

                    }}

                    .total {{
                        font-weight: bold;
                    }}

                </style>

                </head>
                <body onload="window.print()">
                <button onclick="window.print()">🖨 Print</button>

                <div class="container">
                    {all_bills_html}
                </div>

                </body>
                </html>
                """

        tmp = tempfile.NamedTemporaryFile(
            delete=False, suffix=".html", mode="w", encoding="utf-8"
        )
        tmp.write(full_html)
        tmp.close()

        webbrowser.open(f"file://{tmp.name}")

    def detailed_print_selected(self):
        selected_ids = []

        for rid, data in self._ui_map.items():
            if rid == "_summary":
                continue
            if data.get("selected") and data["selected"].get():
                selected_ids.append(rid)

        if not selected_ids:
            messagebox.showwarning("No selection", "Select tenants first")
            return

        all_bills_html = ""

        for record_id in selected_ids:

            idx = self._id_index.get(record_id)
            if idx is None:
                continue

            row = self.df.iloc[idx]

            credit = self.clean_credit(row.get("Credit", ""))
            credit_text = f" ({credit})" if credit else ""

            bill_html = f"""
            <div class="bill-page">

            <div class="bill-header">
                <div class="meta">
                    <div><strong>Date:</strong> {datetime.now().strftime("%Y-%m-%d %I:%M %p")}</div>
                </div>
            </div>

            <div class="meta">
                <strong>Block:</strong> {row['Block']}<br>
                <strong>Name:</strong> {row['Name']}<br>
                <strong>Room:</strong> {row['Room']}<br>
                <strong>Month:</strong> {row['Month']} {row['Year']}
            </div>

            <table>
                <tr>
                    <th class="center">Item</th>
                    <th class="center">Units</th>
                    <th class="center">Rate</th>
                    <th class="right">Charge</th>
                </tr>

                <tr><td>Rent</td><td class="center">-</td><td class="center">-</td><td class="right">{self.safe_int(row['Rent'])}</td></tr>
                <tr><td>Water</td><td class="center">{self.safe_int(row['Wtr Units'])}</td><td class="center">{self.safe_int(row['Wtr Rate'])}</td><td class="right">{self.safe_int(row['Wtr Chg'])}</td></tr>
                <tr><td>Electricity</td><td class="center">{self.safe_int(row['Elec Units'])}</td><td class="center">{self.safe_int(row['Elec Rate'])}</td><td class="right">{self.safe_int(row['Elec Chg'])}</td></tr>
                <tr><td>Internet</td><td class="center">{self.safe_int(row['Net Units'])}</td><td class="center">{self.safe_int(row['Net Rate'])}</td><td class="right">{self.safe_int(row['Net Chg'])}</td></tr>
                <tr><td>Misc</td><td class="center">-</td><td class="center">-</td><td class="right">{self.safe_int(row['Misc'])}</td></tr>
                <tr><td>Old Balance</td><td class="center">-</td><td class="center">-</td><td class="right">{self.safe_int(row['Balance'])}</td></tr>

                <tr class="total-row">
                    <td colspan="3">Total Bill</td>
                    <td class="right">{self.safe_int(row['Total'])}</td>
                </tr>
            </table>
            
            
    
            <div class="summary">
                <div><span>Paid</span><span>{self.safe_int(row['Paid'])}{credit_text}</span></div>
                <div><span>Outstanding</span><span>{self.safe_int(row['Outstanding'])}</span></div>
                <div><span>Advance</span><span>{self.safe_int(row['Advance'])}</span></div>
            </div>

            </div>
            """

            all_bills_html += bill_html

        full_html = f"""
                <!doctype html>
                <html>
                <head>
                <meta charset="utf-8">
                <title>All Bills</title>

                <style>
                    @media print {{
                        button {{
                            display: none;
                            }}

                        body {{
                            margin: 10px;
                        }}
                        h2 {{
                            display: none;
                        }}
                    }}

                    body {{
                        font-family: "Segoe UI", Arial;
                        margin: 20px;
                    }}

                    .bill-container {{
                        display: grid;
                        grid-template-columns: repeat(2, 1fr);
                        gap: 10px
                    }}

                    .bill-page {{   
                    
                        border: 1px solid #ccc;
                        padding: 6px;
                        font-size: 10px;
                        line-height: 1.1;
                        min-height: 170px;       /* ✅ allows safe expansion */
                        box-sizing: border-box;
                        page-break-inside: avoid;

                        display: flex;           /* ✅ prevents overlap */
                        flex-direction: column;
                        justify-content: space-between;

                    }}

                    .bill-header {{
                        display: flex;
                        justify-content: space-between;
                    }}

                    table {{
                        width: 100%;
                        border-collapse: collapse;
                        margin-top: 5px;
                    }}

                    th, td {{
                        border: 1px solid black;
                        padding: 2px;
                        font-size: 11px;
                    }}

                    th {{
                        background: #eee;
                    }}

                    .right {{ text-align: right; }}
                    .center {{ text-align: center; }}

                    .total-row {{
                        background: #e0f2fe;
                        font-weight: bold;
                    }}

                    .summary {{
                        margin-top: 3px;
                        font-size: 10px;
                    }}

                    .summary div {{
                        display: flex;
                        justify-content: space-between;
                        font-size: 11px;
                    }}
                    
                    @page {{
                        size: A4;
                        margin: 10mm;
                        
                        .bill-container {{
                            display: grid;
                            grid-template-columns: repeat(2, 1fr);
                            gap: 10px;
                        }}

                    }}

                    </style>


                </head>
                <body onload="window.print()">
                
                <button onclick="window.print()">🖨 Print</button>

                <div class="bill-container">
                    {all_bills_html}
                </div>

                </body>
                </html>
                """

        tmp = tempfile.NamedTemporaryFile(
            delete=False, suffix=".html", mode="w", encoding="utf-8"
        )
        tmp.write(full_html)
        tmp.close()

        webbrowser.open(f"file://{tmp.name}")


if __name__ == "__main__":
    root = tk.Tk()

    # ✅ HIDE ROOT WINDOW
    root.withdraw()

    app = TenantManagerGUI(root)

    # ✅ open dashboard directly
    root.after(100, app.show_all)

    root.mainloop()
