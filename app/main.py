"""
iPhone Media Browser v0.0.5
"""

import os
import sys
import tempfile
import threading
import time
import shutil
import hashlib
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import win32com.client
    import pythoncom
except ImportError:
    sys.exit("❌ pip install pywin32")

try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont
except ImportError:
    sys.exit("❌ pip install Pillow")

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_OK = True
except ImportError:
    HEIC_OK = False

try:
    import cv2
    CV2_OK = True
except ImportError:
    CV2_OK = False

try:
    LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    LANCZOS = Image.LANCZOS

IMG_EXT = frozenset({
    '.jpg', '.jpeg', '.png', '.heic', '.heif', '.dng',
    '.bmp', '.gif', '.tiff', '.tif', '.webp',
    '.cr2', '.nef', '.arw',
})
VID_EXT = frozenset({
    '.mov', '.mp4', '.m4v', '.avi', '.3gp', '.mkv',
})
SIDECAR_EXT = frozenset({'.aae'})
MEDIA_EXT = IMG_EXT | VID_EXT | SIDECAR_EXT
PREVIEW_SIZE = (450, 450)


def fmt_size(n):
    if not n:
        return "—"
    for u in ("Б", "КБ", "МБ", "ГБ"):
        if abs(n) < 1024:
            return f"{n:.1f} {u}" if isinstance(n, float) else f"{n} {u}"
        n /= 1024
    return f"{n:.2f} ТБ"


def safe_date(d):
    if d is None:
        return None
    if isinstance(d, datetime):
        return d
    try:
        return datetime(d.year, d.month, d.day, d.hour, d.minute, d.second)
    except Exception:
        return None


def file_type_label(ext, is_video, is_sidecar):
    if is_sidecar:
        return "📝 AAE"
    return "🎬 Видео" if is_video else "📷 Фото"


def fmt_elapsed(sec):
    if sec < 60:
        return f"{sec:.0f} сек"
    return f"{sec / 60:.1f} мин"


# ══════════════════════════════════════════════
#  MTP
# ══════════════════════════════════════════════
class MTPDevice:
    KW = ("iphone", "ipad", "apple")

    def find_device(self):
        sh = win32com.client.Dispatch("Shell.Application")
        pc = sh.Namespace(17)
        if pc is None:
            return None, None
        for item in pc.Items():
            if any(k in item.Name.lower() for k in self.KW):
                return item.GetFolder, item.Name
        return None, None

    def list_children(self, folder, cb=None):
        folders = []
        files = 0
        try:
            items = folder.Items()
            for item in items:
                try:
                    nm = item.Name
                    if cb:
                        cb(nm)
                    if item.IsFolder:
                        folders.append((item.GetFolder, nm))
                    else:
                        files += 1
                except Exception:
                    continue
        except Exception:
            pass
        return folders, files

    def scan_folder_recursive(self, folder, prefix, cancel_flag, file_cb=None):
        result = []
        self._walk(folder, prefix, result, cancel_flag, file_cb)
        return result

    def _walk(self, folder, prefix, out, cf, file_cb):
        if cf and cf():
            return
        try:
            items = folder.Items()
        except Exception:
            return
        for item in items:
            if cf and cf():
                return
            try:
                if item.IsFolder:
                    sub = f"{prefix}/{item.Name}" if prefix else item.Name
                    self._walk(item.GetFolder, sub, out, cf, file_cb)
                else:
                    name = item.Name
                    ext = os.path.splitext(name)[1].lower()
                    if ext not in MEDIA_EXT:
                        continue
                    sz = 0
                    try:
                        sz = int(item.Size)
                    except Exception:
                        pass
                    dt = None
                    try:
                        dt = safe_date(item.ModifyDate)
                    except Exception:
                        pass
                    fpath = f"{prefix}/{name}" if prefix else name
                    entry = {
                        "name": name, "path": fpath, "size": sz,
                        "date": dt, "ext": ext,
                        "is_video": ext in VID_EXT,
                        "is_sidecar": ext in SIDECAR_EXT,
                    }
                    out.append(entry)
                    if file_cb:
                        file_cb(len(out), name)
            except Exception:
                continue

    def copy_file(self, mtp_path, dest_dir):
        pythoncom.CoInitialize()
        try:
            sh = win32com.client.Dispatch("Shell.Application")
            pc = sh.Namespace(17)
            if pc is None:
                return None
            device = None
            for it in pc.Items():
                if any(k in it.Name.lower() for k in self.KW):
                    device = it
                    break
            if device is None:
                return None
            cur = device.GetFolder
            parts = mtp_path.replace("\\", "/").split("/")
            filename = parts[-1]
            for fn in parts[:-1]:
                found = False
                for it in cur.Items():
                    if it.IsFolder and it.Name == fn:
                        cur = it.GetFolder
                        found = True
                        break
                if not found:
                    return None
            target = None
            for it in cur.Items():
                if not it.IsFolder and it.Name == filename:
                    target = it
                    break
            if target is None:
                return None
            os.makedirs(dest_dir, exist_ok=True)
            dst_ns = sh.Namespace(dest_dir)
            if dst_ns is None:
                return None
            dst_ns.CopyHere(target, 4 | 16 | 512 | 1024)
            local_path = os.path.join(dest_dir, filename)
            for _ in range(600):
                if os.path.exists(local_path) and os.path.getsize(local_path) > 0:
                    time.sleep(0.3)
                    return local_path
                time.sleep(0.1)
            return local_path if os.path.exists(local_path) else None
        except Exception:
            return None
        finally:
            pythoncom.CoUninitialize()


# ══════════════════════════════════════════════
#  GUI
# ══════════════════════════════════════════════
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("📱 iPhone Media Browser v0.0.5")
        self.root.geometry("1350x780")
        self.root.minsize(900, 500)
        self.mtp = MTPDevice()
        self.all_files = []
        self.shown = []
        self.tmp_dir = tempfile.mkdtemp(prefix="iphone_browser_")
        self.sort_key = None
        self.sort_rev = False
        self._pv_photo = None
        self._loading_lock = threading.Lock()
        self._scan_cancel = False
        self._timer_running = False
        self._timer_start = 0
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._quit)

    def _build_ui(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=28)
        style.configure("G.Horizontal.TProgressbar",
                        troughcolor="#e0e0e0", background="#4caf50")

        # Toolbar
        tb = ttk.Frame(self.root)
        tb.pack(fill="x", padx=5, pady=5)
        self.btn_connect = ttk.Button(tb, text="📱 Подключить iPhone",
                                      command=self._on_connect)
        self.btn_connect.pack(side="left", padx=3)
        self.btn_cancel = ttk.Button(tb, text="⏹ Отмена",
                                     command=self._cancel, state="disabled")
        self.btn_cancel.pack(side="left", padx=3)
        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=8)
        self.btn_sz = ttk.Button(tb, text="⇅ Размер",
                                 command=lambda: self._sort("size"),
                                 state="disabled")
        self.btn_sz.pack(side="left", padx=3)
        self.btn_dt = ttk.Button(tb, text="⇅ Дата",
                                 command=lambda: self._sort("date"),
                                 state="disabled")
        self.btn_dt.pack(side="left", padx=3)
        ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Label(tb, text="Тип:").pack(side="left")
        self.filter_var = tk.StringVar(value="Все")
        cb = ttk.Combobox(tb, textvariable=self.filter_var,
                          values=["Все", "Фото", "Видео", "AAE"],
                          state="readonly", width=7)
        cb.pack(side="left", padx=3)
        cb.bind("<<ComboboxSelected>>", lambda _: self._apply_filter())
        ttk.Label(tb, text="  Поиск:").pack(side="left")
        self.search_var = tk.StringVar()
        se = ttk.Entry(tb, textvariable=self.search_var, width=20)
        se.pack(side="left", padx=3)
        se.bind("<Return>", lambda _: self._apply_filter())
        ttk.Button(tb, text="🔍", width=3,
                   command=self._apply_filter).pack(side="left")
        self.lbl_count = ttk.Label(tb, text="")
        self.lbl_count.pack(side="right", padx=8)

        # ══════ Прогресс-панель ══════
        self.pf = ttk.LabelFrame(self.root, text="")
        inner = ttk.Frame(self.pf)
        inner.pack(fill="x", padx=10, pady=8)

        # Строка 1: Шаг + таймер
        row_top = ttk.Frame(inner)
        row_top.pack(fill="x")

        self.step_var = tk.StringVar(value="")
        ttk.Label(row_top, textvariable=self.step_var,
                  font=("Segoe UI", 12, "bold")).pack(side="left")

        self.timer_var = tk.StringVar(value="")
        ttk.Label(row_top, textvariable=self.timer_var,
                  font=("Segoe UI", 12), foreground="#666").pack(side="right")

        # Строка 2: прогресс-бар + процент
        row_bar = ttk.Frame(inner)
        row_bar.pack(fill="x", pady=(6, 0))

        self.pct_var = tk.StringVar(value="")
        self.pct_lbl = ttk.Label(row_bar, textvariable=self.pct_var,
                                 font=("Segoe UI", 16, "bold"), width=6)
        self.pct_lbl.pack(side="left")

        self.pbar = ttk.Progressbar(row_bar, length=500,
                                    style="G.Horizontal.TProgressbar")
        self.pbar.pack(side="left", fill="x", expand=True, padx=(8, 0))

        # Строка 3: что делаем
        self.action_var = tk.StringVar(value="")
        ttk.Label(inner, textvariable=self.action_var,
                  font=("Segoe UI", 10)).pack(anchor="w", pady=(6, 0))

        # Строка 4: текущий элемент
        self.current_var = tk.StringVar(value="")
        ttk.Label(inner, textvariable=self.current_var,
                  foreground="#888", wraplength=900).pack(anchor="w", pady=(2, 0))

        # Строка 5: ETA (только для шага 3)
        self.eta_var = tk.StringVar(value="")
        ttk.Label(inner, textvariable=self.eta_var,
                  foreground="#4caf50",
                  font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(2, 0))

        # ══════ PanedWindow ══════
        pw = ttk.PanedWindow(self.root, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        lf = ttk.Frame(pw)
        pw.add(lf, weight=3)
        cols = ("name", "size", "date", "type", "ext", "folder")
        self.tree = ttk.Treeview(lf, columns=cols, show="headings",
                                 selectmode="browse")
        for cid, text, w in [("name", "Имя файла", 220),
                              ("size", "Размер", 90),
                              ("date", "Дата", 145),
                              ("type", "Тип", 75),
                              ("ext", "Расш.", 55),
                              ("folder", "Папка", 240)]:
            self.tree.heading(cid, text=text,
                              command=lambda c=cid: self._sort_col(c))
            self.tree.column(cid, width=w, minwidth=40)
        vsb = ttk.Scrollbar(lf, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(lf, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        lf.rowconfigure(0, weight=1)
        lf.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", lambda _: self._open_selected())

        rf = ttk.LabelFrame(pw, text="Предпросмотр")
        pw.add(rf, weight=1)
        self.pv_label = ttk.Label(rf, text="Выберите файл",
                                  anchor="center", justify="center")
        self.pv_label.pack(fill="both", expand=True, padx=8, pady=8)
        self.info_label = ttk.Label(rf, text="", wraplength=340, justify="left")
        self.info_label.pack(fill="x", padx=8, pady=4)
        bf = ttk.Frame(rf)
        bf.pack(pady=6)
        self.btn_open = ttk.Button(bf, text="▶ Открыть",
                                   command=self._open_selected, state="disabled")
        self.btn_open.pack(side="left", padx=6)
        self.btn_save = ttk.Button(bf, text="💾 Сохранить",
                                   command=self._save_selected, state="disabled")
        self.btn_save.pack(side="left", padx=6)

        self.status_var = tk.StringVar(value="Нажмите «Подключить iPhone»")
        ttk.Label(self.root, textvariable=self.status_var,
                  relief="sunken", anchor="w"
                  ).pack(fill="x", side="bottom", padx=5, pady=2)

    # ── Прогресс-хелперы ──
    def _show_pf(self):
        children = self.root.pack_slaves()
        pw_list = [w for w in children if isinstance(w, ttk.PanedWindow)]
        if pw_list:
            self.pf.pack(fill="x", padx=5, pady=(0, 5), before=pw_list[0])
        else:
            self.pf.pack(fill="x", padx=5, pady=(0, 5))

    def _hide_pf(self):
        self.pf.pack_forget()

    def _start_timer(self):
        self._timer_start = time.time()
        self._timer_running = True
        self._tick_timer()

    def _tick_timer(self):
        if not self._timer_running:
            return
        elapsed = time.time() - self._timer_start
        self.timer_var.set(f"⏱ {fmt_elapsed(elapsed)}")
        self.root.after(500, self._tick_timer)

    def _stop_timer(self):
        self._timer_running = False

    def _bar_indeterminate(self):
        """Переключить полосу в режим бегущей анимации."""
        self.pbar.config(mode="indeterminate")
        self.pbar.start(15)
        self.pct_var.set("⏳")

    def _bar_determinate(self, pct=0):
        """Переключить полосу в режим конкретного %."""
        self.pbar.stop()
        self.pbar.config(mode="determinate")
        self.pbar["value"] = pct
        self.pct_var.set(f"{pct:.0f}%")

    def _bar_done(self):
        self.pbar.stop()
        self.pbar.config(mode="determinate")
        self.pbar["value"] = 100
        self.pct_var.set("✅")

    # Потокобезопасные обновления
    def _ui(self, func):
        """Выполнить func в главном потоке."""
        self.root.after(0, func)

    def _set_step(self, step_text):
        self._ui(lambda: self.step_var.set(step_text))
        self._ui(lambda: self.pf.config(text=step_text))

    def _set_action(self, text):
        self._ui(lambda: self.action_var.set(text))

    def _set_current(self, text):
        if len(text) > 90:
            text = "…" + text[-87:]
        self._ui(lambda: self.current_var.set(f"📂 {text}"))

    def _set_eta(self, text):
        self._ui(lambda: self.eta_var.set(text))

    # ── Подключение ──
    def _on_connect(self):
        self._scan_cancel = False
        self.btn_connect["state"] = "disabled"
        self.btn_cancel["state"] = "normal"
        self._show_pf()
        self.step_var.set("Запуск…")
        self.action_var.set("")
        self.current_var.set("")
        self.eta_var.set("")
        self._bar_indeterminate()
        self._start_timer()
        threading.Thread(target=self._scan_thread, daemon=True).start()

    def _cancel(self):
        self._scan_cancel = True

    def _scan_thread(self):
        try:
            pythoncom.CoInitialize()

            # ═══════════════════════════════════
            #  ШАГ 1 из 3: Поиск iPhone
            # ═══════════════════════════════════
            self._set_step("ШАГ 1 из 3 — Поиск iPhone")
            self._set_action("Ищу подключённое устройство Apple…")
            self._ui(self._bar_indeterminate)

            folder, dev_name = self.mtp.find_device()
            if folder is None:
                self._ui(lambda: messagebox.showerror(
                    "iPhone не найден",
                    "• iPhone подключён?\n"
                    "• Экран разблокирован?\n"
                    "• «Доверять этому компьютеру»?\n"
                    "• iTunes установлен?"))
                self._ui(self._scan_done)
                return

            if self._scan_cancel:
                self._ui(self._scan_done)
                return

            self._set_action(f"✅ Найден: {dev_name}")

            # ═══════════════════════════════════
            #  ШАГ 2 из 3: Определение структуры
            # ═══════════════════════════════════
            self._set_step("ШАГ 2 из 3 — Чтение структуры папок")
            self._set_action("Читаю корневую папку…")
            self._set_current("")

            # Уровень 1: "Internal Storage"
            storages, _ = self.mtp.list_children(
                folder,
                cb=lambda n: (
                    self._set_action(f"Корень → «{n}»"),
                    self._set_current(n)
                ))

            if self._scan_cancel:
                self._ui(self._scan_done)
                return

            if not storages:
                self._ui(lambda: messagebox.showwarning(
                    "Пусто", "Не найдено хранилищ. Разблокируйте iPhone."))
                self._ui(self._scan_done)
                return

            # Уровень 2: "DCIM" и т.д.
            level2 = []
            for st_folder, st_name in storages:
                if self._scan_cancel:
                    break
                self._set_action(f"Читаю «{st_name}»…")
                self._set_current(st_name)

                subs, _ = self.mtp.list_children(
                    st_folder,
                    cb=lambda n, sn=st_name: (
                        self._set_action(f"«{sn}» → нашёл «{n}»"),
                        self._set_current(f"{sn}/{n}")
                    ))
                for sf, sn in subs:
                    level2.append((sf, f"{st_name}/{sn}"))

            if self._scan_cancel:
                self._ui(self._scan_done)
                return

            # Уровень 3: "100APPLE", "101APPLE", …
            work_units = []
            for l2_folder, l2_path in level2:
                if self._scan_cancel:
                    break
                self._set_action(f"Читаю «{l2_path}»…")
                self._set_current(l2_path)

                children, _ = self.mtp.list_children(
                    l2_folder,
                    cb=lambda n, p=l2_path: (
                        self._set_action(f"«{p}» → «{n}»"),
                        self._set_current(f"{p}/{n}")
                    ))

                if children:
                    for ch_f, ch_n in children:
                        work_units.append((ch_f, f"{l2_path}/{ch_n}"))
                else:
                    work_units.append((l2_folder, l2_path))

            if self._scan_cancel:
                self._ui(self._scan_done)
                return

            total_units = len(work_units)
            if total_units == 0:
                self._ui(lambda: messagebox.showwarning("Пусто", "Папки пусты."))
                self._ui(self._scan_done)
                return

            self._set_action(f"✅ Найдено {total_units} папок для сканирования")

            # ═══════════════════════════════════
            #  ШАГ 3 из 3: Сканирование файлов
            # ═══════════════════════════════════
            self._set_step("ШАГ 3 из 3 — Чтение файлов")
            self._ui(lambda: self._bar_determinate(0))

            all_files = []
            t_scan = time.time()

            for i, (wu_folder, wu_path) in enumerate(work_units):
                if self._scan_cancel:
                    break

                pct = i / total_units * 100
                elapsed = time.time() - t_scan

                if i > 0 and pct > 0:
                    eta = elapsed / pct * (100 - pct)
                    eta_str = f"⏱ Осталось: ~{fmt_elapsed(eta)}"
                else:
                    eta_str = "⏱ вычисляю…"

                n_files = len(all_files)
                folder_label = wu_path.split("/")[-1] if "/" in wu_path else wu_path

                self._ui(lambda p=pct: self._bar_determinate(p))
                self._set_action(
                    f"Папка {i + 1} из {total_units}  «{folder_label}»  "
                    f"│  Найдено: {n_files} файлов")
                self._set_current(wu_path)
                self._set_eta(eta_str)

                # Колбэк на каждый найденный файл — обновляем счётчик
                last_update = [time.time()]

                def on_file(count, fname, _idx=i):
                    now = time.time()
                    if now - last_update[0] > 0.3:  # не чаще 3 раз/сек
                        last_update[0] = now
                        self._set_action(
                            f"Папка {_idx + 1}/{total_units}  │  "
                            f"Найдено: {len(all_files) + count} файлов  │  "
                            f"Последний: {fname}")

                files = self.mtp.scan_folder_recursive(
                    wu_folder, wu_path,
                    cancel_flag=lambda: self._scan_cancel,
                    file_cb=on_file)
                all_files.extend(files)

            # ═══════════════════════════════════
            #  ГОТОВО
            # ═══════════════════════════════════
            total_sec = time.time() - self._timer_start
            self.all_files = all_files

            self._ui(self._bar_done)
            self._set_step(f"✅ ГОТОВО — {len(all_files)} файлов")
            self._set_action(
                f"Время: {fmt_elapsed(total_sec)}  │  "
                f"Папок: {total_units}")
            self._set_current("")
            self._set_eta("")

            self._ui(self._after_scan)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self._ui(lambda: messagebox.showerror("Ошибка", str(e)))
        finally:
            self._ui(self._scan_done)
            pythoncom.CoUninitialize()

    def _scan_done(self):
        self._stop_timer()
        self.btn_connect["state"] = "normal"
        self.btn_cancel["state"] = "disabled"
        try:
            self.pbar.stop()
        except Exception:
            pass
        self.root.after(6000, self._hide_pf)

    def _after_scan(self):
        photos = sum(1 for f in self.all_files
                     if not f["is_video"] and not f["is_sidecar"])
        videos = sum(1 for f in self.all_files if f["is_video"])
        aae = sum(1 for f in self.all_files if f["is_sidecar"])
        total_sz = sum(f.get("size", 0) for f in self.all_files)
        self.status_var.set(
            f"Всего: {len(self.all_files)} ({fmt_size(total_sz)})  │  "
            f"📷 {photos}  🎬 {videos}  📝 {aae}")
        self.btn_sz["state"] = self.btn_dt["state"] = "normal"
        self._apply_filter()

        ext_counts = {}
        for f in self.all_files:
            e = f["ext"].upper()
            ext_counts[e] = ext_counts.get(e, 0) + 1
        print("\n📊 " + "  ".join(
            f"{e}:{c}" for e, c in
            sorted(ext_counts.items(), key=lambda x: -x[1])))

    # ── Фильтр / сортировка ──
    def _apply_filter(self):
        ft = self.filter_var.get()
        q = self.search_var.get().lower().strip()
        out = []
        for f in self.all_files:
            if ft == "Фото" and (f["is_video"] or f["is_sidecar"]):
                continue
            if ft == "Видео" and not f["is_video"]:
                continue
            if ft == "AAE" and not f["is_sidecar"]:
                continue
            if q and q not in f["name"].lower() and q not in f["path"].lower():
                continue
            out.append(f)
        self.shown = out
        self._fill_tree()

    def _fill_tree(self):
        self.tree.delete(*self.tree.get_children())
        for i, f in enumerate(self.shown):
            sz = fmt_size(f["size"])
            dt = ""
            if f["date"]:
                try:
                    dt = f["date"].strftime("%Y-%m-%d  %H:%M")
                except Exception:
                    pass
            tp = file_type_label(f["ext"], f["is_video"], f["is_sidecar"])
            folder = "/".join(f["path"].split("/")[:-1]) or "/"
            self.tree.insert("", "end", iid=str(i),
                             values=(f["name"], sz, dt, tp,
                                     f["ext"].upper(), folder))
        total_sz = sum(f.get("size", 0) for f in self.shown)
        self.lbl_count.config(
            text=f"Показано: {len(self.shown)} ({fmt_size(total_sz)})")

    def _sort(self, key):
        if self.sort_key == key:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_key = key
            self.sort_rev = False
        if key == "size":
            self.shown.sort(key=lambda x: x.get("size") or 0,
                            reverse=self.sort_rev)
            self.btn_sz.config(text=f"Размер {'↓' if self.sort_rev else '↑'}")
            self.btn_dt.config(text="⇅ Дата")
        else:
            self.shown.sort(key=lambda x: x.get("date") or datetime.min,
                            reverse=self.sort_rev)
            self.btn_dt.config(text=f"Дата {'↓' if self.sort_rev else '↑'}")
            self.btn_sz.config(text="⇅ Размер")
        self._fill_tree()

    def _sort_col(self, col):
        m = {
            "name":   ("name",   lambda x: x["name"].lower()),
            "size":   ("size",   lambda x: x.get("size") or 0),
            "date":   ("date",   lambda x: x.get("date") or datetime.min),
            "type":   ("type",   lambda x: (x["is_sidecar"], x["is_video"])),
            "ext":    ("ext",    lambda x: x["ext"]),
            "folder": ("folder", lambda x: x["path"].lower()),
        }
        if col not in m:
            return
        kn, kf = m[col]
        if self.sort_key == kn:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_key = kn
            self.sort_rev = False
        self.shown.sort(key=kf, reverse=self.sort_rev)
        self._fill_tree()

    # ── Превью ──
    def _get_local(self, f):
        h = hashlib.md5(f["path"].encode()).hexdigest()[:10]
        sub = os.path.join(self.tmp_dir, h)
        local = os.path.join(sub, f["name"])
        if os.path.exists(local) and os.path.getsize(local) > 0:
            return local
        return self.mtp.copy_file(f["path"], sub)

    def _selected_file(self):
        sel = self.tree.selection()
        if not sel:
            return None
        idx = int(sel[0])
        return self.shown[idx] if 0 <= idx < len(self.shown) else None

    def _on_select(self, _):
        f = self._selected_file()
        if not f:
            return
        self.btn_open["state"] = self.btn_save["state"] = "normal"
        info = [
            f"📄 {f['name']}",
            f"📏 {fmt_size(f['size'])}",
            f"📅 {f['date'].strftime('%Y-%m-%d %H:%M') if f['date'] else '—'}",
            f"📁 {f['path']}",
            f"🏷  {file_type_label(f['ext'], f['is_video'], f['is_sidecar'])} ({f['ext']})",
        ]
        if f["is_sidecar"]:
            info.append("\n💡 AAE — данные редактирования iOS")
        self.info_label.config(text="\n".join(info))
        self.pv_label.config(image="", text="⏳ Загрузка…")
        threading.Thread(target=self._pv_thread, args=(f,), daemon=True).start()

    def _pv_thread(self, f):
        if not self._loading_lock.acquire(blocking=False):
            return
        try:
            pythoncom.CoInitialize()
            local = self._get_local(f)
            if not local or not os.path.exists(local):
                self._ui(lambda: self.pv_label.config(
                    image="", text="❌ Не удалось загрузить"))
                return
            f["_local"] = local
            if f["is_sidecar"]:
                self._pv_aae(local, f)
            elif f["is_video"]:
                self._pv_video(local, f)
            else:
                self._pv_image(local, f)
        except Exception as ex:
            self._ui(lambda: self.pv_label.config(image="", text=f"❌ {ex}"))
        finally:
            pythoncom.CoUninitialize()
            self._loading_lock.release()

    def _pv_image(self, path, f):
        try:
            img = Image.open(path)
            try:
                from PIL import ImageOps
                img = ImageOps.exif_transpose(img)
            except Exception:
                pass
            orig = img.size
            img.thumbnail(PREVIEW_SIZE, LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self._ui(lambda: self._set_pv(photo))
            self._ui(lambda: self.status_var.set(
                f"📷 {f['name']} ({orig[0]}×{orig[1]})"))
        except Exception as ex:
            h = "\n💡 pip install pillow-heif" if (
                f["ext"] in (".heic", ".heif") and not HEIC_OK) else ""
            self._ui(lambda: self.pv_label.config(image="", text=f"❌ {ex}{h}"))

    def _pv_video(self, path, f):
        if CV2_OK:
            try:
                cap = cv2.VideoCapture(path)
                total = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
                fps = cap.get(cv2.CAP_PROP_FPS) or 30
                w_ = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH))
                h_ = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
                dur = total / fps if fps else 0
                if total > 10:
                    cap.set(cv2.CAP_PROP_POS_FRAMES, total // 10)
                ok, frame = cap.read()
                cap.release()
                if ok:
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    img = Image.fromarray(frame)
                    img.thumbnail(PREVIEW_SIZE, LANCZOS)
                    draw = ImageDraw.Draw(img, "RGBA")
                    cx, cy = img.size[0] // 2, img.size[1] // 2
                    draw.ellipse((cx - 30, cy - 30, cx + 30, cy + 30),
                                 fill=(0, 0, 0, 120))
                    draw.polygon([(cx - 10, cy - 15), (cx - 10, cy + 15),
                                  (cx + 15, cy)], fill=(255, 255, 255, 200))
                    photo = ImageTk.PhotoImage(img)
                    ds = f"{int(dur // 60)}:{int(dur % 60):02d}" if dur else ""
                    self._ui(lambda: self._set_pv(photo))
                    self._ui(lambda: self.status_var.set(
                        f"🎬 {f['name']} ({w_}×{h_} {ds}) 2×клик▶"))
                    return
            except Exception:
                pass
        self._ui(lambda: self.pv_label.config(
            image="", text=f"🎬 {f['name']}\n\n2×клик ▶"))

    def _pv_aae(self, path, f):
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as fh:
                txt = fh.read(1500)
            self._ui(lambda: self.pv_label.config(
                image="", text=f"📝 AAE\n{'─' * 35}\n{txt[:800]}"))
        except Exception as ex:
            self._ui(lambda: self.pv_label.config(image="", text=f"❌ {ex}"))

    def _set_pv(self, photo):
        self._pv_photo = photo
        self.pv_label.config(image=photo, text="")

    # ── Открыть / Сохранить ──
    def _open_selected(self):
        f = self._selected_file()
        if not f:
            return

        def w():
            try:
                pythoncom.CoInitialize()
                loc = f.get("_local") or self._get_local(f)
                if loc and os.path.exists(loc):
                    os.startfile(loc)
                else:
                    self._ui(lambda: messagebox.showerror(
                        "Ошибка", "Не удалось скопировать файл."))
            except Exception as ex:
                self._ui(lambda: messagebox.showerror("Ошибка", str(ex)))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=w, daemon=True).start()

    def _save_selected(self):
        f = self._selected_file()
        if not f:
            return
        dest = filedialog.asksaveasfilename(
            initialfile=f["name"], defaultextension=f["ext"],
            filetypes=[("Все", "*.*")])
        if not dest:
            return

        def w():
            try:
                pythoncom.CoInitialize()
                loc = f.get("_local") or self._get_local(f)
                if loc and os.path.exists(loc):
                    shutil.copy2(loc, dest)
                    self._ui(lambda: messagebox.showinfo(
                        "Готово", f"Сохранено:\n{dest}"))
                else:
                    self._ui(lambda: messagebox.showerror(
                        "Ошибка", "Не удалось скопировать."))
            except Exception as ex:
                self._ui(lambda: messagebox.showerror("Ошибка", str(ex)))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=w, daemon=True).start()

    def _quit(self):
        self._scan_cancel = True
        self._stop_timer()
        try:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
        except Exception:
            pass
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    print("=" * 55)
    print("  📱 iPhone Media Browser v0.0.5")
    print(f"  HEIC: {'✅' if HEIC_OK else '❌ pip install pillow-heif'}")
    print(f"  Video: {'✅' if CV2_OK else '❌ pip install opencv-python'}")
    print("=" * 55)
    app = App()
    app.run()