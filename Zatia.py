import tkinter as tk
from tkinter import messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap import ttk
from PIL import Image, ImageTk
import sqlite3
import os
import shutil
import tempfile
import webbrowser
import pandas as pd
from datetime import datetime, timedelta
import sys

# External windows (must exist in same folder)
from medical_referral_window import MedicalReferralWindow
from visit_card_window import VisitCardWindow
from departure_card_window import DepartureCardWindow
from work_card_window import WorkCardWindow

DB_NAME = "MNKalajey.db"
PHOTO_DIR = "photos"

if not os.path.exists(PHOTO_DIR):
    os.makedirs(PHOTO_DIR)

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS elements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    father_name TEXT,
    mother_name TEXT,
    birth_info TEXT,
    specialty TEXT,
    arrival_date TEXT,
    nfoos TEXT,
    battalion TEXT,
    camp TEXT,
    photo TEXT,
    last_transfer TEXT,
    previous_camp TEXT,
    previous_battalion TEXT
)
""")
# Ensure new columns exist if older DB missing them
for col in ["last_transfer", "previous_camp", "previous_battalion"]:
    try:
        cursor.execute(f"ALTER TABLE elements ADD COLUMN {col} TEXT")
    except sqlite3.OperationalError:
        pass
conn.commit()

CAMPS = ["معسكر اليرموك", "معسكر أسود الصحراء"]

def save_photo(source_path):
    """
    - If no path provided -> return empty string.
    - If user selected a photo that is already inside PHOTO_DIR -> reuse same path (do not duplicate).
    - Otherwise copy into PHOTO_DIR, avoiding name collisions.
    """
    if not source_path:
        return ""
    try:
        source_abs = os.path.abspath(source_path)
    except Exception:
        return ""
    photo_dir_abs = os.path.abspath(PHOTO_DIR)
    # If the source is already inside the photos directory, just return it (reuse same file).
    try:
        if os.path.commonpath([source_abs, photo_dir_abs]) == photo_dir_abs:
            return source_path
    except Exception:
        # In case of differing drives on Windows, fallback to simple check
        if os.path.dirname(source_abs) == photo_dir_abs:
            return source_path
    if not os.path.isfile(source_path):
        return ""
    filename = os.path.basename(source_path)
    dest_path = os.path.join(PHOTO_DIR, filename)
    base, ext = os.path.splitext(filename)
    counter = 1
    while os.path.exists(dest_path):
        filename = f"{base}_{counter}{ext}"
        dest_path = os.path.join(PHOTO_DIR, filename)
        counter += 1
    try:
        shutil.copy2(source_path, dest_path)
    except Exception:
        shutil.copy(source_path, dest_path)
    return dest_path

class MNKalajeyApp:
    def __init__(self, root):
        self.root = root

        # Use ttkbootstrap Style and a modern theme (default dark)
        self.style = tb.Style(theme="darkly")

        # Window sizing (static)
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        win_w = int(self.screen_width * 0.83)
        win_h = int(self.screen_height * 0.87)
        self.root.geometry(f"{win_w}x{win_h}")
        self.root.title("MNKalajey - ذاتية المعسكرات والكتائب")

        # Try to apply style background, fallback to white
        try:
            self.root.configure(bg=self.style.colors.bg)
        except Exception:
            self.root.configure(bg="#FFFFFF")

        # Fixed default font size and family (no automatic scaling on resize)
        # Requirement: start with font size 2 points smaller than previous default
        # Previous default used earlier was 11 -> start with 9
        self.base_font_size = 11
        initial_size = max(8, self.base_font_size - 2)  # start two points smaller, enforce minimum 8
        self.font_family = "Segoe UI"
        self.font_size = initial_size
        self.update_fonts(self.font_size)

        # button sizing (static-ish)
        self.small_btn_width = max(10, int(self.screen_width / 160))
        self.small_btn_height = max(1, int(self.screen_height / 300))

        self.selected_photo_path = None
        self.selected_element_id = None
        self.sort_order = tk.StringVar(value="تصاعدي")

        # Build UI
        self.create_top_buttons()
        self.create_main_area()
        self.create_footer()
        self.load_camps()
        self.load_battalions()
        self.load_elements()

        # Note: No automatic font resizing bound to <Configure> — per request.

    def _apply_font_recursive(self, widget):
        """
        Walk widget subtree and attempt to apply font config where supported.
        """
        try:
            widget.configure(font=self.mid_font_bold)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._apply_font_recursive(child)

    def update_fonts(self, size):
        """
        Apply a manual font size across the app (called only by user buttons or explicit calls).
        """
        self.font_size = size
        self.mid_font = (self.font_family, self.font_size)
        self.mid_font_bold = (self.font_family, max(8, self.font_size + 1), "bold")

        try:
            self.style.configure("TLabel", font=self.mid_font_bold)
            self.style.configure("TButton", font=self.mid_font_bold)
            self.style.configure("TEntry", font=self.mid_font)
            self.style.configure("TCombobox", font=self.mid_font)
            self.style.configure("Treeview", font=self.mid_font)
            self.style.configure("Treeview.Heading", font=self.mid_font_bold)
        except Exception:
            pass

        # Apply to existing widgets
        try:
            self._apply_font_recursive(self.root)
        except Exception:
            pass

        # Also apply to stored references (safer)
        if hasattr(self, "statistics_label"):
            try:
                self.statistics_label.config(font=self.mid_font_bold)
            except Exception:
                pass
        if hasattr(self, "entry_btns"):
            for btn in self.entry_btns:
                try:
                    btn.config(width=self.small_btn_width, font=self.mid_font_bold)
                except Exception:
                    pass
        if hasattr(self, "top_btns"):
            for btn in self.top_btns:
                try:
                    btn.config(width=self.small_btn_width, font=self.mid_font_bold)
                except Exception:
                    pass
        if hasattr(self, "footer_label"):
            try:
                self.footer_label.config(font=self.mid_font_bold)
            except Exception:
                pass
        # Reconfigure tree style once more
        try:
            self.style.configure("Treeview", font=self.mid_font)
            self.style.configure("Treeview.Heading", font=self.mid_font_bold)
        except Exception:
            pass

    def set_theme(self, theme_name):
        """
        Switch theme using ttkbootstrap and re-apply fonts.
        """
        try:
            self.style.theme_use(theme_name)
        except Exception:
            try:
                self.style = tb.Style(theme=theme_name)
            except Exception:
                messagebox.showerror("خطأ", f"تعذر تغيير الثيم إلى: {theme_name}")
                return
        try:
            self.root.configure(bg=self.style.colors.bg)
        except Exception:
            pass
        self.update_fonts(self.font_size)

    def adjust_font_size(self, delta):
        """
        Manual font size change (via buttons). Bounds enforced.
        """
        new_size = max(8, min(36, self.font_size + delta))
        if new_size != self.font_size:
            self.update_fonts(new_size)

    def create_top_buttons(self):
        top_frame = ttk.Frame(self.root)
        top_frame.pack(side=tk.TOP, fill=tk.X, padx=50, pady=(7, 0))

        # Left controls: theme and font size
        left_ctrl = ttk.Frame(top_frame)
        left_ctrl.pack(side=tk.LEFT, padx=4)

        self.theme_light_btn = ttk.Button(left_ctrl, text="فاتح", command=lambda: self.set_theme("flatly"), bootstyle="outline-success")
        self.theme_light_btn.pack(side=tk.LEFT, padx=4)
        self.theme_dark_btn = ttk.Button(left_ctrl, text="داكن", command=lambda: self.set_theme("darkly"), bootstyle="outline-success")
        self.theme_dark_btn.pack(side=tk.LEFT, padx=4)

        self.font_dec_btn = ttk.Button(left_ctrl, text="تصغير الخط", command=lambda: self.adjust_font_size(-1), bootstyle="warning")
        self.font_dec_btn.pack(side=tk.LEFT, padx=6)
        self.font_inc_btn = ttk.Button(left_ctrl, text="تكبير الخط", command=lambda: self.adjust_font_size(1), bootstyle="warning")
        self.font_inc_btn.pack(side=tk.LEFT, padx=4)

        # Right-side action buttons
        self.top_btns = []
        def make_btn(text, cmd, bootstyle="info"):
            btn = ttk.Button(top_frame, text=text, command=cmd, bootstyle=bootstyle)
            btn.pack(side=tk.RIGHT, padx=6, ipadx=4, ipady=2)
            self.top_btns.append(btn)
            return btn

        make_btn("النقل فترة محددة", self.print_transferred_between_dates, "secondary")
        make_btn("نقل معسكرات", self.move_between_camps, "secondary")
        make_btn("بطاقة زيارة", self.open_visit_card, "info")
        make_btn("أمر مهمة", self.open_work_card, "info")
        make_btn("أمر نقل", self.open_departure_card, "info")
        make_btn("إحالة طبية", self.open_medical_referral, "info")
        make_btn("إحصائية الكتيبة", self.show_statistics_per_battalion, "secondary")
        make_btn("طباعة العدد كامل", self.print_full_statistics, "success")
        make_btn("العدد كامل", self.show_statistics_popup, "secondary")
        make_btn("طباعة بالتصنيف", self.print_sorted, "success")

    def create_main_area(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=6)

        input_frame = ttk.Frame(main_frame)
        input_frame.pack(side=tk.TOP, fill=tk.X)

        entries_frame = ttk.Labelframe(input_frame, text="إضافة / تعديل عنصر")
        entries_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,5), pady=0, ipadx=7, ipady=7)

        labels = [
            "الاسم الكامل:",
            "اسم الأب:",
            "اسم الأم:",
            "مكان وتاريخ الولادة:",
            "الاختصاص:",
            "تاريخ الوصول (يوم/شهر/سنة):",
            "النفوس:",
            "اسم الكتيبة:"
        ]
        # Create entries dictionary keyed by exact label text (including colon)
        self.entries = {}
        for i, txt in enumerate(labels):
            lbl = ttk.Label(entries_frame, text=txt)
            lbl.grid(row=i, column=0, sticky=tk.W, pady=3, padx=8)
            if txt == "تاريخ الوصول (يوم/شهر/سنة):":
                date_frame = ttk.Frame(entries_frame)
                date_frame.grid(row=i, column=1, pady=3, padx=8, sticky=tk.W)
                days = [f"{d:02d}" for d in range(1, 32)]
                self.arrival_day = ttk.Combobox(date_frame, values=days, width=4, state="readonly")
                self.arrival_day.pack(side=tk.LEFT, padx=(0,3))
                months = [f"{m:02d}" for m in range(1, 13)]
                self.arrival_month = ttk.Combobox(date_frame, values=months, width=4, state="readonly")
                self.arrival_month.pack(side=tk.LEFT, padx=(0,3))
                # Year range: raise upper bound to 2035 (or current_year+1 if that is greater)
                current_year = datetime.now().year
                max_year = max(current_year + 1, 2035)
                years = [str(y) for y in range(max_year, 1949, -1)]
                self.arrival_year = ttk.Combobox(date_frame, values=years, width=7, state="readonly")
                self.arrival_year.pack(side=tk.LEFT, padx=(0,3))
                # set reasonable defaults (if available)
                try:
                    today = datetime.now()
                    self.arrival_year.set(str(today.year))
                    self.arrival_month.set(f"{today.month:02d}")
                    self.arrival_day.set(f"{today.day:02d}")
                except Exception:
                    pass

                # Holder to keep same interface as Entry.get()
                class _Holder:
                    def get(inner):
                        try:
                            y = self.arrival_year.get()
                            m = self.arrival_month.get()
                            d = self.arrival_day.get()
                            if y and m and d:
                                return f"{y}-{m}-{d}"
                        except Exception:
                            pass
                        return ""
                self.entries[txt] = _Holder()
            else:
                ent = ttk.Entry(entries_frame, width=28)
                ent.grid(row=i, column=1, pady=3, padx=8)
                self.entries[txt] = ent

        ttk.Label(entries_frame, text="اسم المعسكر:").grid(row=8, column=0, sticky=tk.W, pady=3, padx=8)
        self.camp_combobox = ttk.Combobox(entries_frame, state="readonly", width=26)
        self.camp_combobox.grid(row=8, column=1, pady=3, padx=8)

        # Photo display (tk.Label to support image)
        self.photo_label = tk.Label(entries_frame, relief="solid", borderwidth=3, background="#444444")
        self.photo_label.grid(row=0, column=2, rowspan=7, padx=10, pady=6)
        photo_btn = ttk.Button(entries_frame, text="اختر صورة شخصية", command=self.select_photo, bootstyle="secondary")
        photo_btn.grid(row=8, column=2, padx=10, pady=6)

        excel_btn_frame = ttk.Frame(entries_frame)
        excel_btn_frame.grid(row=8, column=4, padx=10, pady=3)
        export_excel_btn = ttk.Button(excel_btn_frame, text="Excel تصدير ملف", command=self.export_all_to_excel, bootstyle="success")
        export_excel_btn.pack(side=tk.LEFT, padx=3)
        import_excel_btn = ttk.Button(excel_btn_frame, text="Excel استيراد ملف", command=self.import_excel, bootstyle="primary")
        import_excel_btn.pack(side=tk.LEFT, padx=3)

        # Right-side entry action buttons
        right_btn_frame = ttk.Frame(input_frame)
        right_btn_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(6,3), pady=8)
        self.entry_btns = []
        def make_entry_btn(text, cmd, bootstyle="primary"):
            btn = ttk.Button(right_btn_frame, text=text, command=cmd, bootstyle=bootstyle)
            btn.pack(side=tk.RIGHT, padx=6, pady=5)
            self.entry_btns.append(btn)
            return btn

        make_entry_btn("إضافة عنصر", self.add_element, "success")
        make_entry_btn("تحديث عنصر", self.update_element, "info")
        make_entry_btn("حذف عنصر", self.delete_element, "danger")
        make_entry_btn("تفريغ المربعات", self.clear_inputs, "secondary")

        # Bottom: list and filters
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=0, pady=2)

        list_frame = ttk.Labelframe(bottom_frame, text="عناصر الكتائب والمعسكرات")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        filter_frame = ttk.Frame(list_frame)
        filter_frame.pack(fill=tk.X, padx=10, pady=7)

        ttk.Label(filter_frame, text="اختيار المعسكر:").pack(side=tk.LEFT, padx=5)
        self.camp_filter = ttk.Combobox(filter_frame, state="readonly", width=11)
        self.camp_filter.pack(side=tk.LEFT, padx=5)
        self.camp_filter.bind("<<ComboboxSelected>>", lambda e: self.load_battalion_filter())

        ttk.Label(filter_frame, text="اختيار كتيبة:").pack(side=tk.LEFT, padx=5)
        self.battalion_filter = ttk.Combobox(filter_frame, state="readonly", width=11)
        self.battalion_filter.pack(side=tk.LEFT, padx=5)
        self.battalion_filter.bind("<<ComboboxSelected>>", lambda e: self.load_elements())

        self.sort_columns = [
            "id", "name", "father_name", "mother_name",
            "birth_info", "specialty", "arrival_date", "nfoos", "battalion", "camp", "photo", "last_transfer"
        ]
        self.sort_labels = [
            "الرقم", "الاسم الكامل", "اسم الأب", "اسم الأم",
            "مكان وتاريخ الولادة", "الاختصاص", "تاريخ الوصول", "النفوس",
            "اسم الكتيبة", "اسم المعسكر", "مسار الصورة", "تاريخ النقل"
        ]

        ttk.Label(filter_frame, text="ترتيب حسب:").pack(side=tk.LEFT, padx=5)
        self.sort_by = ttk.Combobox(filter_frame, state="readonly", values=self.sort_labels, width=11)
        try:
            self.sort_by.current(1)
        except Exception:
            pass
        self.sort_by.pack(side=tk.LEFT, padx=5)
        self.sort_by.bind("<<ComboboxSelected>>", lambda e: self.load_elements())

        ttk.Label(filter_frame, text="ترتيب:").pack(side=tk.LEFT, padx=5)
        self.sort_order_cb = ttk.Combobox(filter_frame, state="readonly", values=["تصاعدي", "تنازلي"], width=9, textvariable=self.sort_order)
        try:
            self.sort_order_cb.current(0)
        except Exception:
            pass
        self.sort_order_cb.pack(side=tk.LEFT, padx=5)
        self.sort_order_cb.bind("<<ComboboxSelected>>", lambda e: self.load_elements())

        ttk.Label(filter_frame, text="بحث شامل:").pack(side=tk.LEFT, padx=5)
        self.search_entry = ttk.Entry(filter_frame, width=19)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind("<KeyRelease>", lambda e: self.load_elements())

        self.columns = [
            "id", "name", "father_name", "mother_name",
            "birth_info", "specialty",
            "arrival_date", "nfoos", "battalion", "camp", "photo", "last_transfer"
        ]
        headings = [
            "#", "الاسم الكامل", "اسم الأب", "اسم الأم",
            "مكان وتاريخ الولادة", "الاختصاص",
            "تاريخ الوصول", "النفوس", "اسم الكتيبة", "اسم المعسكر", "مسار الصورة", "تاريخ النقل"
        ]
        self.tree = ttk.Treeview(list_frame, columns=self.columns, show="headings", selectmode="browse")
        for col, head in zip(self.columns, headings):
            self.tree.heading(col, text=head)
            self.tree.column(col, width=90 if col not in ["photo", "birth_info", "specialty"] else 140, anchor=tk.CENTER)
        self.tree.column("id", width=40, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=7)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)

        self.statistics_label = ttk.Label(self.root, text="", anchor=tk.W)
        self.statistics_label.pack(fill=tk.X, padx=12, pady=(0, 7))

    def create_footer(self):
        self.footer_label = ttk.Label(self.root,
            text="تمت البرمجة عام 2025م من قبل محمد نجيب قلعه جي /أبو نجيب/ إدلب - حاس",
            anchor="center"
        )
        self.footer_label.pack(side=tk.BOTTOM, fill=tk.X, pady=3)

    def select_photo(self):
        path = filedialog.askopenfilename(
            title="اختر صورة شخصية",
            filetypes=[("صور", "*.jpg *.jpeg *.png *.bmp"), ("كل الملفات", "*.*")]
        )
        if path:
            self.selected_photo_path = path
            self.show_photo(path)

    def show_photo(self, path):
        try:
            img = Image.open(path).resize((200, 200))
            self.photo_img = ImageTk.PhotoImage(img)
            self.photo_label.config(image=self.photo_img)
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر عرض الصورة: {e}")

    def get_arrival_date_str(self):
        # Prefer comboboxes; return YYYY-MM-DD or empty
        y = self.arrival_year.get() if hasattr(self, "arrival_year") else ""
        m = self.arrival_month.get() if hasattr(self, "arrival_month") else ""
        d = self.arrival_day.get() if hasattr(self, "arrival_day") else ""
        if y and m and d:
            return f"{y}-{m}-{d}"
        try:
            return self.entries["تاريخ الوصول (يوم/شهر/سنة):"].get().strip()
        except Exception:
            return ""

    def get_input_data(self):
        # Use keys exactly as created above
        name = self.entries["الاسم الكامل:"].get().strip()
        father = self.entries["اسم الأب:"].get().strip()
        mother = self.entries["اسم الأم:"].get().strip()
        birth_info = self.entries["مكان وتاريخ الولادة:"].get().strip()
        specialty = self.entries["الاختصاص:"].get().strip()
        arrival = self.get_arrival_date_str().strip()
        nfoos = self.entries["النفوس:"].get().strip()
        battalion = self.entries["اسم الكتيبة:"].get().strip()
        camp = self.camp_combobox.get().strip()
        if not name or not battalion or not camp:
            messagebox.showerror("خطأ", "الاسم الكامل والكتيبة والمعسكر مطلوبان.")
            return None
        if arrival:
            try:
                datetime.strptime(arrival, "%Y-%m-%d")
            except Exception:
                messagebox.showerror("خطأ", "تاريخ الوصول يجب أن يكون بالصورة: YYYY-MM-DD (اختر اليوم/الشهر/السنة).")
                return None
        return (name, father, mother, birth_info, specialty, arrival, nfoos, battalion, camp)

    def clear_inputs(self):
        for key, entry in self.entries.items():
            try:
                entry.delete(0, tk.END)
            except Exception:
                pass
        try:
            self.photo_label.config(image="", bg="#444444")
        except Exception:
            self.photo_label.config(image="")
        self.selected_photo_path = None
        self.selected_element_id = None
        try:
            self.camp_combobox.set("")
        except Exception:
            pass
        if hasattr(self, "arrival_day"):
            self.arrival_day.set("")
        if hasattr(self, "arrival_month"):
            self.arrival_month.set("")
        if hasattr(self, "arrival_year"):
            self.arrival_year.set("")

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        values = item.get('values', [])
        # Values: [id, name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos, battalion, camp, photo, last_transfer]
        # Map fields carefully to avoid index errors
        field_keys = [
            "الاسم الكامل:", "اسم الأب:", "اسم الأم:",
            "مكان وتاريخ الولادة:", "الاختصاص:",
            "تاريخ الوصول (يوم/شهر/سنة):", "النفوس:", "اسم الكتيبة:"
        ]
        for idx, key in enumerate(field_keys):
            try:
                if key == "تاريخ الوصول (يوم/شهر/سنة):":
                    arrival_val = values[6] if len(values) > 6 else ""
                    if arrival_val:
                        parsed = None
                        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
                            try:
                                parsed = datetime.strptime(arrival_val, fmt)
                                break
                            except Exception:
                                continue
                        if parsed:
                            if hasattr(self, "arrival_day"):
                                self.arrival_day.set(f"{parsed.day:02d}")
                            if hasattr(self, "arrival_month"):
                                self.arrival_month.set(f"{parsed.month:02d}")
                            if hasattr(self, "arrival_year"):
                                # ensure year value exists in combobox values before set
                                y = str(parsed.year)
                                try:
                                    if y in self.arrival_year['values']:
                                        self.arrival_year.set(y)
                                    else:
                                        # if year not in list (e.g., parsed year > max_year), insert temporarily
                                        vals = list(self.arrival_year['values'])
                                        vals.insert(0, y)
                                        self.arrival_year['values'] = vals
                                        self.arrival_year.set(y)
                                except Exception:
                                    try:
                                        self.arrival_year.set(y)
                                    except Exception:
                                        pass
                        else:
                            if hasattr(self, "arrival_day"):
                                self.arrival_day.set("")
                            if hasattr(self, "arrival_month"):
                                self.arrival_month.set("")
                            if hasattr(self, "arrival_year"):
                                self.arrival_year.set("")
                    else:
                        if hasattr(self, "arrival_day"):
                            self.arrival_day.set("")
                        if hasattr(self, "arrival_month"):
                            self.arrival_month.set("")
                        if hasattr(self, "arrival_year"):
                            self.arrival_year.set("")
                else:
                    # normal entry fields: map to values[idx+1] (since values[0] is id)
                    val = values[idx+1] if len(values) > idx+1 else ""
                    try:
                        self.entries[key].delete(0, tk.END)
                        self.entries[key].insert(0, val if val is not None else "")
                    except Exception:
                        pass
            except Exception:
                pass

        # camp is at index 9
        try:
            camp_val = values[9] if len(values) > 9 else ""
            self.camp_combobox.set(camp_val if camp_val is not None else "")
        except Exception:
            pass

        # photo path is at index 10
        try:
            photo_val = values[10] if len(values) > 10 else ""
            if photo_val:
                self.show_photo(photo_val)
                self.selected_photo_path = photo_val
            else:
                try:
                    self.photo_label.config(image="", bg="#444444")
                except Exception:
                    self.photo_label.config(image="")
                self.selected_photo_path = None
        except Exception:
            pass

        # id is at index 0
        try:
            self.selected_element_id = values[0] if len(values) > 0 else None
        except Exception:
            self.selected_element_id = None

    def add_element(self):
        data = self.get_input_data()
        if not data:
            return
        photo_path = save_photo(self.selected_photo_path)
        cursor.execute("""
            INSERT INTO elements (name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos, battalion, camp, photo, last_transfer, previous_camp, previous_battalion)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, "", "", "")
        """, (*data, photo_path))
        conn.commit()
        messagebox.showinfo("نجاح", "تمت الإضافة بنجاح.")
        self.load_battalions()
        self.load_elements()
        self.clear_inputs()

    def update_element(self):
        if not self.selected_element_id:
            messagebox.showerror("خطأ", "اختر عنصرًا للتحديث.")
            return
        data = self.get_input_data()
        if not data:
            return
        photo_path = ""
        if self.selected_photo_path:
            photo_path = save_photo(self.selected_photo_path)
        if photo_path:
            cursor.execute("""
                UPDATE elements SET name=?, father_name=?, mother_name=?, birth_info=?, specialty=?, arrival_date=?, nfoos=?, battalion=?, camp=?, photo=?
                WHERE id=?
            """, (*data, photo_path, self.selected_element_id))
        else:
            cursor.execute("""
                UPDATE elements SET name=?, father_name=?, mother_name=?, birth_info=?, specialty=?, arrival_date=?, nfoos=?, battalion=?, camp=?
                WHERE id=?
            """, (*data, self.selected_element_id))
        conn.commit()
        messagebox.showinfo("نجاح", "تم التحديث بنجاح.")
        self.load_battalions()
        self.load_elements()
        self.clear_inputs()

    def delete_element(self):
        if not self.selected_element_id:
            messagebox.showerror("خطأ", "اختر عنصرًا للحذف.")
            return
        if messagebox.askyesno("تأكيد", "هل أنت متأكد؟"):
            cursor.execute("DELETE FROM elements WHERE id=?", (self.selected_element_id,))
            conn.commit()
            self.load_battalions()
            self.load_elements()
            self.clear_inputs()

    def move_between_camps(self):
        if not self.selected_element_id:
            messagebox.showerror("خطأ", "اختر عنصرًا للنقل أولاً من الجدول.")
            return
        cursor.execute("SELECT camp, battalion FROM elements WHERE id=?", (self.selected_element_id,))
        row = cursor.fetchone()
        if not row:
            messagebox.showerror("خطأ", "العنصر غير موجود!")
            return
        current_camp, current_battalion = row
        other_camps = [camp for camp in CAMPS if camp != current_camp]
        if not other_camps:
            messagebox.showerror("خطأ", "لا يوجد معسكر آخر للنقل إليه.")
            return
        to_camp = other_camps[0]
        if messagebox.askyesno("تأكيد النقل", f"هل تريد نقل العنصر إلى {to_camp}؟"):
            nowstr = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            cursor.execute(
                "UPDATE elements SET previous_camp=?, camp=?, last_transfer=?, previous_battalion=? WHERE id=?",
                (current_camp, to_camp, nowstr, current_battalion, self.selected_element_id)
            )
            conn.commit()
            messagebox.showinfo("تم النقل", f"تم نقل العنصر إلى {to_camp}")
            self.load_battalions()
            self.load_elements()
            self.clear_inputs()

    def load_camps(self):
        camps = list(CAMPS)
        self.camp_combobox['values'] = camps
        try:
            self.camp_filter['values'] = ["الكل"] + camps
            self.camp_combobox.set("")
            self.camp_filter.current(0)
        except Exception:
            pass

    def load_battalions(self):
        camp = self.camp_filter.get()
        if camp == "الكل" or not camp:
            cursor.execute("SELECT DISTINCT battalion FROM elements WHERE battalion IS NOT NULL AND battalion != ''")
        else:
            cursor.execute("SELECT DISTINCT battalion FROM elements WHERE battalion IS NOT NULL AND battalion != '' AND camp=?", (camp,))
        battalions = [row[0] for row in cursor.fetchall()]
        battalions = sorted(battalions)
        battalions.insert(0, "الكل")
        self.battalion_filter['values'] = battalions
        try:
            self.battalion_filter.current(0)
        except Exception:
            pass
        self.load_elements()

    def load_battalion_filter(self):
        # same as load_battalions but does not refresh other controls
        camp = self.camp_filter.get()
        if camp == "الكل" or not camp:
            cursor.execute("SELECT DISTINCT battalion FROM elements WHERE battalion IS NOT NULL AND battalion != ''")
        else:
            cursor.execute("SELECT DISTINCT battalion FROM elements WHERE battalion IS NOT NULL AND battalion != '' AND camp=?", (camp,))
        battalions = [row[0] for row in cursor.fetchall()]
        battalions = sorted(battalions)
        battalions.insert(0, "الكل")
        self.battalion_filter['values'] = battalions
        try:
            self.battalion_filter.current(0)
        except Exception:
            pass
        self.load_elements()

    def load_elements(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        camp = self.camp_filter.get()
        battalion = self.battalion_filter.get()
        sort_label = self.sort_by.get()
        sort_key = self.sort_columns[self.sort_labels.index(sort_label)] if sort_label in self.sort_labels else "name"
        sort_order = "ASC" if self.sort_order.get() == "تصاعدي" else "DESC"
        search_text = self.search_entry.get().strip()
        query = f"SELECT id, name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos, battalion, camp, photo, last_transfer FROM elements"
        params = ()
        conditions = []
        if camp != "الكل" and camp:
            conditions.append("camp = ?")
            params += (camp,)
        if battalion != "الكل" and battalion:
            conditions.append("battalion = ?")
            params += (battalion,)
        if search_text:
            search_cols = self.columns
            like_text = f"%{search_text}%"
            conds = [f"CAST({col} AS TEXT) LIKE ?" for col in search_cols]
            conditions.append("(" + " OR ".join(conds) + ")")
            params += tuple([like_text]*len(search_cols))
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        query += f" ORDER BY {sort_key} {sort_order}"
        cursor.execute(query, params)
        rows = cursor.fetchall()
        for r in rows:
            self.tree.insert("", tk.END, values=r)

    def show_statistics_popup(self):
        camp = self.camp_filter.get()
        if camp == "الكل" or not camp:
            cursor.execute("SELECT camp, battalion, COUNT(*) FROM elements GROUP BY camp, battalion")
            stats = cursor.fetchall()
            total = sum(row[2] for row in stats)
            msg = ""
            camps = set(row[0] for row in stats)
            for campx in camps:
                msg += f"المعسكر: {campx}\n"
                camp_stats = [row for row in stats if row[0] == campx]
                for cs in camp_stats:
                    msg += f"   الكتيبة: {cs[1]} - عدد العناصر: {cs[2]}\n"
            msg += f"\nالمجموع الكلي: {total}"
        else:
            cursor.execute("SELECT battalion, COUNT(*) FROM elements WHERE camp=? GROUP BY battalion", (camp,))
            stats = cursor.fetchall()
            total = sum(row[1] for row in stats)
            msg = f"المعسكر: {camp}\n"
            for cs in stats:
                msg += f"   الكتيبة: {cs[0]} - عدد العناصر: {cs[1]}\n"
            msg += f"\nالمجموع الكلي: {total}"
        messagebox.showinfo("إحصائية المعسكرات والكتائب", msg)

    def print_full_statistics(self):
        camp = self.camp_filter.get()
        if camp == "الكل" or not camp:
            cursor.execute("SELECT camp, battalion, COUNT(*) FROM elements GROUP BY camp, battalion")
            stats = cursor.fetchall()
            camps = set(row[0] for row in stats)
            total = sum(row[2] for row in stats)
        else:
            cursor.execute("SELECT camp, battalion, COUNT(*) FROM elements WHERE camp=? GROUP BY camp, battalion", (camp,))
            stats = cursor.fetchall()
            camps = set([camp])
            total = sum(row[2] for row in stats)
        html = """
<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<title>إحصائية كاملة للمعسكرات والكتائب</title>
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #fff; color: #000; }
    table { width: 80%; border-collapse: collapse; margin: 20px auto;}
    th, td { padding: 8px; border: 1px solid #444; text-align: center; background: white; }
    h2 { color: #2979FF; text-align: center;}
</style>
</head>
<body>
<h2>إحصائية كاملة للمعسكرات والكتائب</h2>
"""
        for campx in camps:
            html += f"<h3>المعسكر: {campx}</h3><table><tr><th>الكتيبة</th><th>عدد العناصر</th></tr>"
            camp_stats = [row for row in stats if row[0] == campx]
            for cs in camp_stats:
                html += f"<tr><td>{cs[1]}</td><td>{cs[2]}</td></tr>"
            html += "</table>"
        html += f"<h3 style='text-align:center;'>المجموع الكلي: {total}</h3>"
        html += "</body></html>"
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8")
        tmp_file.write(html)
        tmp_file.close()
        if sys.platform == "win32":
            os.startfile(tmp_file.name)
        else:
            webbrowser.open(f"file://{tmp_file.name}")

    def show_statistics_per_battalion(self):
        battalion = self.battalion_filter.get()
        camp = self.camp_filter.get()
        if battalion == "الكل" or not battalion:
            messagebox.showerror("خطأ", "يرجى اختيار كتيبة معينة لعرض إحصائياتها.")
            return
        if camp == "الكل" or not camp:
            cursor.execute("SELECT COUNT(*) FROM elements WHERE battalion=?", (battalion,))
        else:
            cursor.execute("SELECT COUNT(*) FROM elements WHERE battalion=? AND camp=?", (battalion, camp))
        count = cursor.fetchone()[0]
        messagebox.showinfo(f"إحصائية الكتيبة {battalion}", f"عدد العناصر: {count}")

    def print_battalion(self):
        battalion = self.battalion_filter.get()
        camp = self.camp_filter.get()
        if battalion == "الكل" or not battalion:
            messagebox.showerror("خطأ", "يرجى اختيار كتيبة معينة للطباعة.")
            return
        if camp == "الكل" or not camp:
            cursor.execute("""
                SELECT name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos
                FROM elements WHERE battalion=?
            """, (battalion,))
        else:
            cursor.execute("""
                SELECT name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos
                FROM elements WHERE battalion=? AND camp=?
            """, (battalion, camp))
        rows = cursor.fetchall()
        html = f"""
<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<title>عناصر الكتيبة: {battalion} - المعسكر: {camp}</title>
<style>
    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #fff; color: #000; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ padding: 8px; border: 1px solid #444; text-align: center; background: white; }}
    h2 {{ color: #2979FF; }}
</style>
</head>
<body>
<h2>عناصر الكتيبة: {battalion} - المعسكر: {camp}</h2>
<table>
<tr>
<th>الاسم الكامل</th><th>اسم الأب</th><th>اسم الأم</th>
<th>مكان وتاريخ الولادة</th><th>الاختصاص</th><th>تاريخ الوصول</th><th>النفوس</th>
</tr>
"""
        for r in rows:
            html += "<tr>" + "".join(f"<td>{str(v) if v else ''}</td>" for v in r) + "</tr>"
        html += "</table></body></html>"

        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8")
        tmp_file.write(html)
        tmp_file.close()
        if sys.platform == "win32":
            os.startfile(tmp_file.name)
        else:
            webbrowser.open(f"file://{tmp_file.name}")

    def print_transferred_between_dates(self):
        def do_print():
            # Build from_date and to_date strings from comboboxes and append times
            y1 = from_year.get()
            m1 = from_month.get()
            d1 = from_day.get()
            y2 = to_year.get()
            m2 = to_month.get()
            d2 = to_day.get()
            if not (y1 and m1 and d1 and y2 and m2 and d2):
                messagebox.showerror("خطأ", "��رجى اختيار كامل تاريخي البداية والنهاية (اليوم/الشهر/السنة).")
                return
            try:
                from_dt = datetime.strptime(f"{y1}-{m1}-{d1} 00:00:00", "%Y-%m-%d %H:%M:%S")
                to_dt = datetime.strptime(f"{y2}-{m2}-{d2} 23:59:59", "%Y-%m-%d %H:%M:%S")
            except Exception:
                messagebox.showerror("خطأ", "تواريخ غير صالحة. تأكد من اختيار اليوم/الشهر/السنة بشكل صحيح.")
                return
            cursor.execute("""
            SELECT name, father_name, mother_name, birth_info, specialty, arrival_date,
                   previous_battalion, battalion, previous_camp, camp, last_transfer
            FROM elements WHERE last_transfer IS NOT NULL AND last_transfer >= ? AND last_transfer <= ?
            """, (from_dt.strftime("%Y-%m-%d %H:%M:%S"), to_dt.strftime("%Y-%m-%d %H:%M:%S")))
            rows = cursor.fetchall()
            if not rows:
                messagebox.showinfo("لا يوجد بيانات", "لا يوجد عناصر تم نقلها ضمن الفترة المحددة.")
                return
            html = f"""
<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<title>العناصر المنقولة ضمن الفترة</title>
<style>
    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #fff; color: #000; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ padding: 8px; border: 1px solid #444; text-align: center; background: white; }}
    h2 {{ color: #2979FF; }}
</style>
</head>
<body>
<h2>العناصر المنقولة من {from_dt.strftime("%Y-%m-%d %H:%M:%S")}</h2>
<h2>إلى {to_dt.strftime("%Y-%m-%d %H:%M:%S")}</h2>
<table>
<tr>
<th>الاسم الكامل</th><th>اسم الأب</th><th>اسم الأم</th>
<th>مكان وتاريخ الولادة</th><th>الاختصاص</th><th>تاريخ الوصول</th>
<th>الكتيبة السابقة</th><th>الكتيبة الحالية</th>
<th>المعسكر السابق</th><th>المعسكر الحالي</th>
<th>تاريخ النقل</th>
</tr>
"""
            for r in rows:
                html += "<tr>" + "".join(f"<td>{str(x) if x else ''}</td>" for x in r) + "</tr>"
            html += "</table></body></html>"
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8")
            tmp_file.write(html)
            tmp_file.close()
            if sys.platform == "win32":
                os.startfile(tmp_file.name)
            else:
                webbrowser.open(f"file://{tmp_file.name}")
            win.destroy()

        win = tk.Toplevel(self.root)
        win.title("تحديد فترة النقل")
        win.geometry("420x180")

        # Use comboboxes for from and to (day/month/year) similar to arrival widgets
        frame_from = ttk.Frame(win)
        frame_from.pack(pady=(8,4))
        ttk.Label(frame_from, text="من (يوم/شهر/سنة):").pack(side=tk.LEFT, padx=(0,6))
        days = [f"{d:02d}" for d in range(1, 32)]
        months = [f"{m:02d}" for m in range(1, 13)]
        current_year = datetime.now().year
        max_year = max(current_year + 1, 2035)
        years = [str(y) for y in range(max_year, 1949, -1)]

        from_day = ttk.Combobox(frame_from, values=days, width=4, state="readonly")
        from_day.pack(side=tk.LEFT, padx=(0,3))
        from_month = ttk.Combobox(frame_from, values=months, width=4, state="readonly")
        from_month.pack(side=tk.LEFT, padx=(0,3))
        from_year = ttk.Combobox(frame_from, values=years, width=7, state="readonly")
        from_year.pack(side=tk.LEFT, padx=(0,3))

        frame_to = ttk.Frame(win)
        frame_to.pack(pady=(4,6))
        ttk.Label(frame_to, text="إلى (يوم/شهر/سنة):").pack(side=tk.LEFT, padx=(0,12))
        to_day = ttk.Combobox(frame_to, values=days, width=4, state="readonly")
        to_day.pack(side=tk.LEFT, padx=(0,3))
        to_month = ttk.Combobox(frame_to, values=months, width=4, state="readonly")
        to_month.pack(side=tk.LEFT, padx=(0,3))
        to_year = ttk.Combobox(frame_to, values=years, width=7, state="readonly")
        to_year.pack(side=tk.LEFT, padx=(0,3))

        # set sensible defaults: from = yesterday, to = today
        try:
            today = datetime.now().date()
            yesterday = today - timedelta(days=1)
            from_year.set(str(yesterday.year))
            from_month.set(f"{yesterday.month:02d}")
            from_day.set(f"{yesterday.day:02d}")
            to_year.set(str(today.year))
            to_month.set(f"{today.month:02d}")
            to_day.set(f"{today.day:02d}")
        except Exception:
            pass

        ttk.Button(win, text="طباعة", command=do_print, bootstyle="primary").pack(pady=8)

    def export_all_to_excel(self):
        cursor.execute("SELECT name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos, battalion, camp FROM elements")
        rows = cursor.fetchall()
        columns = [
            "name", "father_name", "mother_name",
            "birth_info", "specialty",
            "arrival_date", "nfoos", "battalion", "camp"
        ]
        df = pd.DataFrame(rows, columns=columns)
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel files", "*.xlsx")],
                                            title="حفظ ملف Excel")
        if path:
            try:
                df.to_excel(path, index=False)
                messagebox.showinfo("نجاح", f"تم تصدير البيانات إلى {path}")
            except Exception as e:
                messagebox.showerror("خطأ", f"فشل التصدير: {e}")

    def import_excel(self):
        path = filedialog.askopenfilename(title="اختر ملف Excel",
                                          filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            df = pd.read_excel(path)
            required_cols = [
                "name", "father_name", "mother_name",
                "birth_info", "specialty",
                "arrival_date", "nfoos", "battalion", "camp"
            ]
            if not set(required_cols).issubset(set(df.columns)):
                messagebox.showerror(
                    "خطأ",
                    "ملف Excel يجب أن يحتوي على الأعمدة التالية: name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos, battalion, camp"
                )
                return
            count_added = 0
            for _, row in df.iterrows():
                values = [str(row[col]).strip() if not pd.isna(row[col]) else "" for col in required_cols]
                cursor.execute("""
                INSERT INTO elements (name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos, battalion, camp, photo, last_transfer, previous_camp, previous_battalion)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, "", "", "", "")
                """, values)
                count_added += 1
            conn.commit()
            messagebox.showinfo("نجاح", f"تم إضافة {count_added} عنصر من ملف Excel.")
            self.load_battalions()
            self.load_elements()
        except Exception as e:
            messagebox.showerror("خطأ", f"فشل الاستيراد: {e}")

    def open_medical_referral(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("خطأ", "يرجى اختيار عنصر أولاً من الجدول.")
            return
        element_id = self.tree.item(selected[0])["values"][0]
        MedicalReferralWindow(self.root, conn, cursor, element_id)

    def open_departure_card(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("خطأ", "يرجى اختيار عنصر أولاً من الجدول.")
            return
        element_id = self.tree.item(selected[0])["values"][0]
        DepartureCardWindow(self.root, conn, cursor, element_id)

    def open_visit_card(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("خطأ", "يرجى اختيار عنصر أولاً من الجدول.")
            return
        element_id = self.tree.item(selected[0])["values"][0]
        VisitCardWindow(self.root, conn, cursor, element_id)

    def open_work_card(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showerror("خطأ", "يرجى اختيار عنصر أولاً من الجدول.")
            return
        element_id = self.tree.item(selected[0])["values"][0]
        WorkCardWindow(self.root, conn, cursor, element_id)

    def print_sorted(self):
        camp = self.camp_filter.get()
        battalion = self.battalion_filter.get()
        sort_label = self.sort_by.get()
        sort_key = self.sort_columns[self.sort_labels.index(sort_label)] if sort_label in self.sort_labels else "name"
        sort_order = "ASC" if self.sort_order.get() == "تصاعدي" else "DESC"
        search_text = self.search_entry.get().strip()
        query = "SELECT name, father_name, mother_name, birth_info, specialty, arrival_date, nfoos, battalion, camp FROM elements"
        params = ()
        conditions = []
        if camp != "الكل" and camp:
            conditions.append("camp = ?")
            params += (camp,)
        if battalion != "الكل" and battalion:
            conditions.append("battalion = ?")
            params += (battalion,)
        if search_text:
            search_cols = ["name", "father_name", "mother_name", "birth_info", "specialty", "arrival_date", "nfoos", "battalion", "camp"]
            like_text = f"%{search_text}%"
            conds = [f"CAST({col} AS TEXT) LIKE ?" for col in search_cols]
            conditions.append("(" + " OR ".join(conds) + ")")
            params += tuple([like_text]*len(search_cols))
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        query += f" ORDER BY {sort_key} {sort_order}"
        cursor.execute(query, params)
        rows = cursor.fetchall()

        html = f"""<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<title>طباعة حسب التصنيف والترتيب</title>
<style>
    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #fff; color: #000; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ padding: 8px; border: 1px solid #444; text-align: center; background: white; }}
    h2 {{ color: #2979FF; }}
</style>
</head>
<body>
<h2>طباعة حسب التصنيف ({sort_label}) وبالترتيب {'تصاعدي' if sort_order == 'ASC' else 'تنازلي'}</h2>
<table>
<tr>
<th>الاسم الكامل</th>
<th>اسم الأب</th>
<th>اسم الأم</th>
<th>مكان وتاريخ الولادة</th>
<th>الاختصاص</th>
<th>تاريخ الوصول</th>
<th>النفوس</th>
<th>اسم الكتيبة</th>
<th>اسم المعسكر</th>
</tr>
"""
        for r in rows:
            html += "<tr>" + "".join(f"<td>{str(v) if v else ''}</td>" for v in r) + "</tr>"
        html += "</table></body></html>"

        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".html", mode="w", encoding="utf-8")
        tmp_file.write(html)
        tmp_file.close()
        if sys.platform == "win32":
            os.startfile(tmp_file.name)
        else:
            webbrowser.open(f"file://{tmp_file.name}")

if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = MNKalajeyApp(root)
    root.mainloop()
    conn.close()