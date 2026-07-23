# -*- coding: utf-8 -*-
"""
Meta Lead Ads 客户留言同步选择器

放置位置：ADS/lead_sync/sync_leads_gui.py
用途：用勾选框选择产品，然后调用 sync_leads.py 增量同步。
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent
FORMS_LIBRARY = BASE_DIR / "_cache" / "forms_library.json"
SYNC_SCRIPT = SCRIPT_DIR / "sync_leads.py"

BG = "#f5f7fb"
CARD = "#ffffff"
TEXT = "#111827"
MUTED = "#6b7280"
BORDER = "#e5e7eb"
PURPLE = "#5432EB"
PURPLE_DARK = "#4326c9"
YELLOW = "#FFAD00"
DANGER = "#b91c1c"
SUCCESS = "#047857"
LOG_BG = "#0b1020"
LOG_FG = "#d1d5db"

FONT = "Microsoft YaHei UI"


def normalize_key(value: str) -> str:
    return "".join(ch for ch in str(value).strip().lower() if ch not in [" ", "-", "_"])


def load_products() -> list[str]:
    if not FORMS_LIBRARY.exists():
        return []
    with FORMS_LIBRARY.open("r", encoding="utf-8-sig") as f:
        data = json.load(f)
    found: dict[str, str] = {}
    if isinstance(data, dict):
        for _page_id, products in data.items():
            if not isinstance(products, dict):
                continue
            for product in products.keys():
                text = str(product)
                found.setdefault(normalize_key(text), text)
    return sorted(found.values(), key=lambda x: x.lower())


def python_cmd() -> list[str]:
    exe = sys.executable or "python"
    return [exe]


class LeadSyncGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ADS Lead Sync")
        self.root.geometry("1180x760")
        self.root.minsize(1080, 700)

        self.products = load_products()
        self.product_vars: dict[str, tk.BooleanVar] = {}
        self.product_rows: dict[str, tk.Frame] = {}
        self.full_var = tk.BooleanVar(value=False)
        self.initial_days_var = tk.StringVar(value="")
        self.start_date_var = tk.StringVar(value="")
        self.end_date_var = tk.StringVar(value="")
        self.search_var = tk.StringVar(value="")
        self.selected_count_var = tk.StringVar(value="已选择 0 个产品")
        self.status_var = tk.StringVar(value="Ready")
        self.running = False

        self.setup_style()
        self.build_ui()
        self.refresh_product_rows()
        self.update_selected_count()

    def setup_style(self) -> None:
        self.root.configure(bg=BG)
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TCheckbutton", background=CARD, foreground=TEXT, font=(FONT, 10))
        style.map("TCheckbutton", background=[("active", CARD)])

    def build_ui(self) -> None:
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=22, pady=20)

        self.build_header(outer)

        main = tk.Frame(outer, bg=BG)
        main.pack(fill="both", expand=True, pady=(16, 0))

        left = self.card(main)
        left.pack(side="left", fill="both", expand=True, padx=(0, 14))

        right = self.card(main, width=360)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        self.build_product_card(left)
        self.build_settings_card(right)

        self.build_log_card(outer)

    def build_header(self, parent: tk.Frame) -> None:
        header = tk.Frame(parent, bg=BG)
        header.pack(fill="x")

        title_block = tk.Frame(header, bg=BG)
        title_block.pack(side="left", fill="x", expand=True)

        tk.Label(
            title_block,
            text="同步客户留言",
            bg=BG,
            fg=TEXT,
            font=(FONT, 21, "bold"),
        ).pack(anchor="w")

        tk.Label(
            title_block,
            text="选择产品和时间段后同步，结果每次新建 Excel 保存到 ADS/lead_sync/",
            bg=BG,
            fg=MUTED,
            font=(FONT, 10),
        ).pack(anchor="w", pady=(5, 0))

        badge = tk.Frame(header, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        badge.pack(side="right", padx=(12, 0), ipady=6, ipadx=12)
        tk.Label(badge, text="默认增量同步", bg=CARD, fg=SUCCESS, font=(FONT, 10, "bold")).pack()
        tk.Label(badge, text="节省 API 请求", bg=CARD, fg=MUTED, font=(FONT, 8)).pack(pady=(2, 0))

    def card(self, parent: tk.Frame, width: int | None = None) -> tk.Frame:
        frame = tk.Frame(parent, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        if width:
            frame.configure(width=width)
        return frame

    def section_title(self, parent: tk.Frame, title: str, subtitle: str | None = None) -> None:
        tk.Label(parent, text=title, bg=CARD, fg=TEXT, font=(FONT, 13, "bold")).pack(anchor="w")
        if subtitle:
            tk.Label(parent, text=subtitle, bg=CARD, fg=MUTED, font=(FONT, 9)).pack(anchor="w", pady=(4, 0))

    def build_product_card(self, parent: tk.Frame) -> None:
        inner = tk.Frame(parent, bg=CARD)
        inner.pack(fill="both", expand=True, padx=20, pady=18)

        top = tk.Frame(inner, bg=CARD)
        top.pack(fill="x")
        title = tk.Frame(top, bg=CARD)
        title.pack(side="left", fill="x", expand=True)
        self.section_title(title, "选择产品", "可搜索，也可以全选/反选；不选产品时会同步全部。")

        tk.Label(top, textvariable=self.selected_count_var, bg="#f3f0ff", fg=PURPLE, font=(FONT, 9, "bold"), padx=10, pady=5).pack(side="right")

        search_wrap = tk.Frame(inner, bg="#f9fafb", highlightthickness=1, highlightbackground=BORDER)
        search_wrap.pack(fill="x", pady=(14, 10))
        tk.Label(search_wrap, text="搜索", bg="#f9fafb", fg=MUTED, font=(FONT, 9)).pack(side="left", padx=(12, 6))
        search = tk.Entry(search_wrap, textvariable=self.search_var, relief="flat", bg="#f9fafb", fg=TEXT, font=(FONT, 10))
        search.pack(side="left", fill="x", expand=True, ipady=8, padx=(0, 10))
        self.search_var.trace_add("write", lambda *_: self.refresh_product_rows())

        actions = tk.Frame(inner, bg=CARD)
        actions.pack(fill="x", pady=(0, 12))
        self.ghost_button(actions, "全选", self.select_all).pack(side="left", padx=(0, 8))
        self.ghost_button(actions, "清空", self.clear_all).pack(side="left", padx=(0, 8))
        self.ghost_button(actions, "反选", self.invert_selection).pack(side="left")

        list_shell = tk.Frame(inner, bg=CARD, highlightthickness=1, highlightbackground=BORDER)
        list_shell.pack(fill="both", expand=True)

        canvas = tk.Canvas(list_shell, bg=CARD, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_shell, orient="vertical", command=canvas.yview)
        self.products_frame = tk.Frame(canvas, bg=CARD)
        self.products_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.products_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        scrollbar.pack(side="right", fill="y", padx=(0, 6), pady=8)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def build_settings_card(self, parent: tk.Frame) -> None:
        inner = tk.Frame(parent, bg=CARD)
        inner.pack(fill="both", expand=True, padx=20, pady=18)
        self.section_title(inner, "同步设置", "选择时间段或直接增量同步。")

        date_box = tk.Frame(inner, bg="#f9fafb", highlightthickness=1, highlightbackground=BORDER)
        date_box.pack(fill="x", pady=(16, 12), ipady=8, ipadx=10)
        tk.Label(date_box, text="时间段（可选）", bg="#f9fafb", fg=TEXT, font=(FONT, 10, "bold")).pack(anchor="w")
        tk.Label(
            date_box,
            text="填了时间段就按该范围导出，不受增量状态和已同步记录影响。",
            bg="#f9fafb",
            fg=MUTED,
            font=(FONT, 8),
            wraplength=235,
            justify="left",
        ).pack(anchor="w", pady=(3, 8))

        start_wrap = tk.Frame(date_box, bg="#ffffff", highlightthickness=1, highlightbackground=BORDER)
        start_wrap.pack(fill="x", pady=(0, 8))
        tk.Label(start_wrap, text="开始", bg="#ffffff", fg=MUTED, font=(FONT, 8), width=6).pack(side="left", padx=(8, 2))
        tk.Entry(start_wrap, textvariable=self.start_date_var, relief="flat", bg="#ffffff", fg=TEXT, font=(FONT, 10)).pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 8))

        end_wrap = tk.Frame(date_box, bg="#ffffff", highlightthickness=1, highlightbackground=BORDER)
        end_wrap.pack(fill="x")
        tk.Label(end_wrap, text="结束", bg="#ffffff", fg=MUTED, font=(FONT, 8), width=6).pack(side="left", padx=(8, 2))
        tk.Entry(end_wrap, textvariable=self.end_date_var, relief="flat", bg="#ffffff", fg=TEXT, font=(FONT, 10)).pack(side="left", fill="x", expand=True, ipady=7, padx=(0, 8))

        tk.Label(
            date_box,
            text="格式：2026-07-01。结束日期包含当天；两个都留空则继续自动增量。",
            bg="#f9fafb",
            fg=MUTED,
            font=(FONT, 8),
            wraplength=235,
            justify="left",
        ).pack(anchor="w", pady=(7, 0))

        mode_box = tk.Frame(inner, bg="#f9fafb", highlightthickness=1, highlightbackground=BORDER)
        mode_box.pack(fill="x", pady=(16, 12), ipady=8, ipadx=10)
        ttk.Checkbutton(mode_box, text="全量重拉 --full", variable=self.full_var).pack(anchor="w")
        tk.Label(
            mode_box,
            text="只在排查缺数据时使用，会增加 API 请求。",
            bg="#f9fafb",
            fg=MUTED,
            font=(FONT, 8),
        ).pack(anchor="w", padx=24, pady=(3, 0))

        # Action area: keep the primary button visible, not hidden below settings.
        action_box = tk.Frame(inner, bg=CARD)
        action_box.pack(fill="x", pady=(18, 14))

        self.run_btn = tk.Button(
            action_box,
            text="开始同步",
            command=self.start_sync,
            bg=PURPLE,
            fg="#ffffff",
            activebackground=PURPLE_DARK,
            activeforeground="#ffffff",
            relief="flat",
            font=(FONT, 14, "bold"),
            cursor="hand2",
        )
        self.run_btn.pack(fill="x", ipady=13)

        tip = tk.Frame(inner, bg="#fff7e6", highlightthickness=1, highlightbackground="#fde68a")
        tip.pack(fill="x", pady=(0, 14), ipady=8, ipadx=10)
        tk.Label(tip, text="提示", bg="#fff7e6", fg="#92400e", font=(FONT, 9, "bold")).pack(anchor="w")
        tk.Label(
            tip,
            text="默认只拉新增留言；不选产品会弹窗确认同步全部。首次同步默认回看 30 天。",
            bg="#fff7e6",
            fg="#92400e",
            font=(FONT, 8),
            wraplength=285,
            justify="left",
        ).pack(anchor="w", pady=(3, 0))

        status_box = tk.Frame(inner, bg="#f9fafb", highlightthickness=1, highlightbackground=BORDER)
        status_box.pack(fill="x", ipady=8, ipadx=10)
        tk.Label(status_box, text="状态", bg="#f9fafb", fg=MUTED, font=(FONT, 8)).pack(anchor="w")
        tk.Label(status_box, textvariable=self.status_var, bg="#f9fafb", fg=TEXT, font=(FONT, 9, "bold"), wraplength=235, justify="left").pack(anchor="w", pady=(3, 0))

        spacer = tk.Frame(inner, bg=CARD)
        spacer.pack(fill="both", expand=True)

        tk.Label(
            inner,
            text="输出位置：ADS/lead_sync/leads_时间戳.xlsx",
            bg=CARD,
            fg=MUTED,
            font=(FONT, 8),
        ).pack(anchor="w")

    def build_log_card(self, parent: tk.Frame) -> None:
        log_card = self.card(parent)
        log_card.pack(fill="both", pady=(14, 0))
        header = tk.Frame(log_card, bg=CARD)
        header.pack(fill="x", padx=16, pady=(12, 6))
        tk.Label(header, text="运行日志", bg=CARD, fg=TEXT, font=(FONT, 11, "bold")).pack(side="left")
        self.ghost_button(header, "清空日志", lambda: self.log_box.delete("1.0", "end")).pack(side="right")

        self.log_box = scrolledtext.ScrolledText(
            log_card,
            height=8,
            bg=LOG_BG,
            fg=LOG_FG,
            insertbackground=LOG_FG,
            font=("Consolas", 9),
            relief="flat",
            padx=10,
            pady=8,
        )
        self.log_box.pack(fill="both", expand=False, padx=16, pady=(0, 14))
        self.append_log("Ready. Choose products and click Start Sync.\n")

    def ghost_button(self, parent: tk.Frame, text: str, command) -> tk.Button:
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg="#f3f4f6",
            fg=TEXT,
            activebackground="#e5e7eb",
            activeforeground=TEXT,
            relief="flat",
            font=(FONT, 9),
            cursor="hand2",
            padx=10,
            pady=5,
        )

    def product_row(self, product: str) -> tk.Frame:
        row = tk.Frame(self.products_frame, bg=CARD)
        var = tk.BooleanVar(value=False)
        self.product_vars[product] = var
        cb = ttk.Checkbutton(row, text=product, variable=var, command=self.update_selected_count)
        cb.pack(anchor="w", fill="x", padx=10, pady=7)
        return row

    def refresh_product_rows(self) -> None:
        for child in self.products_frame.winfo_children():
            child.pack_forget()

        query = normalize_key(self.search_var.get())
        visible = []
        for product in self.products:
            if not query or query in normalize_key(product):
                visible.append(product)

        if not self.products:
            tk.Label(
                self.products_frame,
                text="没有读取到表单库。\n请确认 ADS/_cache/forms_library.json 存在。",
                bg=CARD,
                fg=DANGER,
                font=(FONT, 10),
                wraplength=450,
                justify="left",
            ).pack(anchor="w", padx=14, pady=14)
            return

        if not visible:
            tk.Label(self.products_frame, text="没有匹配的产品", bg=CARD, fg=MUTED, font=(FONT, 10)).pack(anchor="w", padx=14, pady=14)
            return

        for product in visible:
            row = self.product_rows.get(product)
            if row is None:
                row = self.product_row(product)
                self.product_rows[product] = row
            row.pack(fill="x")

    def select_all(self) -> None:
        query = normalize_key(self.search_var.get())
        for name, var in self.product_vars.items():
            if not query or query in normalize_key(name):
                var.set(True)
        self.update_selected_count()

    def clear_all(self) -> None:
        query = normalize_key(self.search_var.get())
        for name, var in self.product_vars.items():
            if not query or query in normalize_key(name):
                var.set(False)
        self.update_selected_count()

    def invert_selection(self) -> None:
        query = normalize_key(self.search_var.get())
        for name, var in self.product_vars.items():
            if not query or query in normalize_key(name):
                var.set(not var.get())
        self.update_selected_count()

    def update_selected_count(self) -> None:
        count = len(self.selected_products())
        self.selected_count_var.set(f"已选择 {count} 个产品")

    def selected_products(self) -> list[str]:
        return [name for name, var in self.product_vars.items() if var.get()]

    def append_log(self, text: str) -> None:
        self.log_box.insert("end", text)
        self.log_box.see("end")

    def start_sync(self) -> None:
        if self.running:
            return
        if not SYNC_SCRIPT.exists():
            messagebox.showerror("错误", f"找不到同步脚本：{SYNC_SCRIPT}")
            return

        products = self.selected_products()
        start_date = self.start_date_var.get().strip()
        end_date = self.end_date_var.get().strip()
        date_re = re.compile(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}$")
        if start_date and not date_re.match(start_date):
            messagebox.showerror("错误", "开始日期格式请填写 YYYY-MM-DD，例如 2026-07-01。")
            return
        if end_date and not date_re.match(end_date):
            messagebox.showerror("错误", "结束日期格式请填写 YYYY-MM-DD，例如 2026-07-31。")
            return

        full = self.full_var.get()
        product_text = ",".join(products)
        if full:
            ok = messagebox.askyesno("确认全量同步", "你勾选了全量重拉，会增加 API 请求。确认继续？")
            if not ok:
                return

        if start_date or end_date:
            ok = messagebox.askyesno("确认时间段导出", "本次会按你填写的时间段导出客户留言，并且不受增量状态和已同步记录影响。确认继续？")
            if not ok:
                return

        if not products:
            ok = messagebox.askyesno("确认同步全部", "你没有选择产品，将同步全部产品表单。确认继续？")
            if not ok:
                return

        cmd = python_cmd() + [str(SYNC_SCRIPT)]
        if start_date:
            cmd += ["--start-date", start_date]
        if end_date:
            cmd += ["--end-date", end_date]
        if product_text:
            cmd += ["--products", product_text]
        if full:
            cmd.append("--full")

        self.running = True
        self.run_btn.configure(state="disabled", text="同步中...")
        self.status_var.set("Running...")
        self.append_log("\n> " + " ".join(f'\"{x}\"' if " " in x else x for x in cmd) + "\n")

        threading.Thread(target=self.run_process, args=(cmd,), daemon=True).start()

    def run_process(self, cmd: list[str]) -> None:
        try:
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            process = subprocess.Popen(
                cmd,
                cwd=str(BASE_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            assert process.stdout is not None
            for line in process.stdout:
                self.root.after(0, self.append_log, line)
            code = process.wait()
            self.root.after(0, self.finish_sync, code)
        except Exception as exc:
            self.root.after(0, self.append_log, f"\nERROR: {exc}\n")
            self.root.after(0, self.finish_sync, 1)

    def finish_sync(self, code: int) -> None:
        self.running = False
        self.run_btn.configure(state="normal", text="开始同步")
        if code == 0:
            self.status_var.set("Done. 已生成新的 Excel，保存在 lead_sync 文件夹。")
            self.append_log("\nSync finished. New Excel saved in ADS/lead_sync/.\n")
        else:
            self.status_var.set("Failed. Please check ADS/日志/.")
            self.append_log("\nSync failed. Please check ADS/日志/.\n")


def main() -> None:
    root = tk.Tk()
    LeadSyncGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
