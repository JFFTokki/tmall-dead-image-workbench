from __future__ import annotations

import os
import queue
import sys
import threading
import time
from pathlib import Path
from tkinter import filedialog, messagebox

ROOT = Path(__file__).resolve().parent
VENDOR = ROOT / "vendor" / "python"
if VENDOR.exists():
    sys.path.insert(0, str(VENDOR))

import customtkinter as ctk

from dead_image_workbench_core import (
    DEFAULT_CORRECTION,
    DEFAULT_HAR,
    DEFAULT_INPUT,
    DEFAULT_TEMPLATE,
    OUTPUT_ROOT,
    RunOutputs,
    create_problem_workbook,
    find_first_existing,
    parse_progress,
    quote_arg,
    repair_mojibake,
    safe_stem,
    DETECT_SCRIPT,
    EXPORT_SCRIPT,
)
import subprocess


class ModernDeadImageWorkbench(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("死图判断工作台")
        self.geometry("1220x780")
        self.minsize(1080, 700)

        self.log_queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.current_outputs: RunOutputs | None = None
        self.last_clean_log = ""
        self.nav_buttons: dict[str, ctk.CTkButton] = {}

        default_input = find_first_existing(DEFAULT_INPUT, ROOT / "export_469_20260706_135143_2932072.xlsx")
        self.source_mode = ctk.StringVar(value="excel")
        self.input_xlsx = ctk.StringVar(value=str(default_input if default_input.exists() else ""))
        self.ids_file = ctk.StringVar()
        self.har_file = ctk.StringVar(value=str(DEFAULT_HAR if DEFAULT_HAR.exists() else ""))
        self.output_dir = ctk.StringVar(value=str(OUTPUT_ROOT))
        self.template_xlsx = ctk.StringVar(value=str(DEFAULT_TEMPLATE if DEFAULT_TEMPLATE.exists() else ""))
        self.correction_xlsx = ctk.StringVar(value=str(DEFAULT_CORRECTION if DEFAULT_CORRECTION.exists() else ""))
        self.batch_size = ctk.StringVar(value="100")
        self.llm_enabled = ctk.BooleanVar(value=False)
        self.llm_api_url = ctk.StringVar()
        self.llm_api_key = ctk.StringVar()
        self.llm_model = ctk.StringVar()

        self.total_images = ctk.StringVar(value="-")
        self.problem_count = ctk.StringVar(value="-")
        self.status_text = ctk.StringVar(value="准备就绪")
        self.progress_percent = ctk.StringVar(value="0%")

        self.colors = {
            "bg": "#eef2f6",
            "surface": "#ffffff",
            "soft": "#f8fafc",
            "ink": "#0f172a",
            "muted": "#64748b",
            "line": "#e2e8f0",
            "blue": "#2563eb",
            "blue_dark": "#1d4ed8",
            "nav": "#0b1220",
        }

        self._build_ui()
        self.after(120, self._drain_log_queue)

    def _build_ui(self) -> None:
        self.configure(fg_color=self.colors["bg"])
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._build_sidebar()
        self._build_main()

    def _build_sidebar(self) -> None:
        nav = ctk.CTkFrame(self, width=238, fg_color=self.colors["nav"], corner_radius=0)
        nav.grid(row=0, column=0, sticky="nsew")
        nav.grid_propagate(False)

        ctk.CTkLabel(
            nav,
            text="死图判断\n工作台",
            justify="left",
            font=ctk.CTkFont("Microsoft YaHei UI", 22, "bold"),
            text_color="#f8fafc",
        ).pack(anchor="w", padx=24, pady=(30, 28))

        self._nav_item(nav, "data", "1  数据准备", active=True, command=lambda: self._navigate("data"))
        self._nav_item(nav, "judge", "2  图片判断", command=lambda: self._navigate("judge"))
        self._nav_item(nav, "result", "3  结果交付", command=lambda: self._navigate("result"))
        self._nav_item(nav, "diagnostic", "4  诊断留存", command=lambda: self._navigate("diagnostic"))

        ctk.CTkFrame(nav, fg_color="transparent").pack(fill="both", expand=True)
        hint = ctk.CTkFrame(nav, fg_color="#111827", corner_radius=18)
        hint.pack(fill="x", padx=18, pady=22)
        ctk.CTkLabel(
            hint,
            text="主结果和问题结果给普通用户；日志和检查点自动留存。",
            wraplength=174,
            justify="left",
            text_color="#cbd5e1",
            font=ctk.CTkFont("Microsoft YaHei UI", 12),
        ).pack(anchor="w", padx=14, pady=14)

    def _nav_item(self, parent: ctk.CTkFrame, key: str, text: str, active: bool = False, command=None) -> None:
        button = ctk.CTkButton(
            parent,
            text=text,
            anchor="w",
            command=command,
            height=50,
            corner_radius=12,
            fg_color=self.colors["blue"] if active else "transparent",
            hover_color="#1e3a8a",
            text_color="#ffffff" if active else "#94a3b8",
            font=ctk.CTkFont("Microsoft YaHei UI", 14, "bold" if active else "normal"),
        )
        button.pack(fill="x", padx=16, pady=(0, 8 if active else 4))
        self.nav_buttons[key] = button

    def _set_active_nav(self, active_key: str) -> None:
        for key, button in self.nav_buttons.items():
            active = key == active_key
            button.configure(
                fg_color=self.colors["blue"] if active else "transparent",
                text_color="#ffffff" if active else "#94a3b8",
                font=ctk.CTkFont("Microsoft YaHei UI", 14, "bold" if active else "normal"),
            )

    def _navigate(self, target: str) -> None:
        self._set_active_nav(target)
        if target == "data":
            self.left_scroll._parent_canvas.yview_moveto(0.0)
        elif target == "judge":
            self.left_scroll._parent_canvas.yview_moveto(0.55)
        elif target == "result":
            self.result_card.focus_set()
        elif target == "diagnostic":
            self.result_card.focus_set()

    def _build_main(self) -> None:
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=28, pady=24)
        main.grid_columnconfigure(0, weight=1)
        main.grid_columnconfigure(1, minsize=360)
        main.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(main, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 20))
        ctk.CTkLabel(
            header,
            text="死图质检工作流",
            text_color=self.colors["ink"],
            font=ctk.CTkFont("Microsoft YaHei UI", 30, "bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="导入源数据，实时查看判断进度，最后只交付两张结果表。",
            text_color=self.colors["muted"],
            font=ctk.CTkFont("Microsoft YaHei UI", 14),
        ).pack(anchor="w", pady=(6, 0))

        left = ctk.CTkScrollableFrame(
            main,
            fg_color="transparent",
            scrollbar_button_color="#cbd5e1",
            scrollbar_button_hover_color="#94a3b8",
        )
        self.left_scroll = left
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 22))
        left.grid_columnconfigure(0, weight=1)

        right = ctk.CTkFrame(main, fg_color="transparent")
        right.grid(row=1, column=1, sticky="nsew")

        self._build_source_card(left)
        self._build_options_card(left)
        self._build_llm_card(left)
        self._build_log_card(left)
        self._build_result_card(right)

    def _card(self, parent: ctk.CTkFrame, row: int, sticky: str = "ew") -> ctk.CTkFrame:
        card = ctk.CTkFrame(parent, fg_color=self.colors["surface"], corner_radius=22)
        card.grid(row=row, column=0, sticky=sticky, pady=(0, 16))
        card.grid_columnconfigure(1, weight=1)
        return card

    def _section_title(self, parent: ctk.CTkFrame, title: str, subtitle: str) -> None:
        ctk.CTkLabel(parent, text=title, text_color=self.colors["ink"], font=ctk.CTkFont("Microsoft YaHei UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=22, pady=(20, 0)
        )
        ctk.CTkLabel(parent, text=subtitle, text_color=self.colors["muted"], font=ctk.CTkFont("Microsoft YaHei UI", 12)).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=22, pady=(4, 16)
        )

    def _build_source_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, 0)
        self._section_title(card, "数据来源", "先选择一种入口。已有 Excel 直接判断；商品 ID 会先导出源数据。")
        self._file_row(card, 2, "已有 Excel", self.input_xlsx, self._choose_input_xlsx, radio_value="excel")
        self._file_row(card, 3, "商品 ID 文件", self.ids_file, self._choose_ids_file, radio_value="ids")
        self._file_row(card, 4, "HAR 登录文件", self.har_file, self._choose_har_file)

    def _build_options_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, 1)
        self._section_title(card, "运行设置", "默认使用已有学习模型和模板参考；过程文件只留作诊断。")
        rows = [
            ("输出目录", self.output_dir, self._choose_output_dir),
            ("示例模板", self.template_xlsx, self._choose_template_xlsx),
            ("修正反馈", self.correction_xlsx, self._choose_correction_xlsx),
        ]
        for index, (label, var, action) in enumerate(rows, start=2):
            self._file_row(card, index, label, var, action)

        ctk.CTkLabel(card, text="检查点批量", text_color="#334155", font=ctk.CTkFont("Microsoft YaHei UI", 13)).grid(
            row=5, column=0, sticky="w", padx=(22, 14), pady=(8, 0)
        )
        ctk.CTkEntry(card, textvariable=self.batch_size, width=120, height=38, corner_radius=10).grid(
            row=5, column=1, sticky="w", pady=(8, 0)
        )

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=6, column=0, columnspan=3, sticky="ew", padx=22, pady=(20, 20))
        self.run_button = ctk.CTkButton(
            actions,
            text="开始运行",
            width=132,
            height=42,
            corner_radius=12,
            fg_color=self.colors["blue"],
            hover_color=self.colors["blue_dark"],
            font=ctk.CTkFont("Microsoft YaHei UI", 14, "bold"),
            command=self._start_run,
        )
        self.run_button.pack(side="left")
        ctk.CTkButton(
            actions,
            text="打开输出目录",
            width=142,
            height=42,
            corner_radius=12,
            fg_color="#ffffff",
            hover_color="#f1f5f9",
            border_color=self.colors["line"],
            border_width=1,
            text_color=self.colors["ink"],
            command=lambda: self._open_path(Path(self.output_dir.get())),
        ).pack(side="left", padx=(12, 0))

    def _build_llm_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, 2)
        self._section_title(card, "大模型辅助", "预留多模态大模型 API 入口；不开启时不影响当前本地判断流程。")

        ctk.CTkCheckBox(
            card,
            text="启用大模型辅助判断",
            variable=self.llm_enabled,
            onvalue=True,
            offvalue=False,
            font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold"),
            text_color="#334155",
        ).grid(row=2, column=0, columnspan=3, sticky="w", padx=22, pady=(0, 10))

        self._entry_row(card, 3, "API 地址", self.llm_api_url, placeholder="例如：https://api.example.com/v1/chat/completions")
        self._entry_row(card, 4, "API Key", self.llm_api_key, placeholder="只保存在本次运行窗口中", secret=True)
        self._entry_row(card, 5, "模型名称", self.llm_model, placeholder="例如：gpt-4o / qwen-vl-max / gemini-vision")

    def _entry_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: ctk.StringVar,
        placeholder: str = "",
        secret: bool = False,
    ) -> None:
        ctk.CTkLabel(parent, text=label, text_color="#334155", font=ctk.CTkFont("Microsoft YaHei UI", 13)).grid(
            row=row, column=0, sticky="w", padx=(22, 14), pady=7
        )
        ctk.CTkEntry(
            parent,
            textvariable=variable,
            placeholder_text=placeholder,
            show="*" if secret else "",
            height=38,
            corner_radius=10,
            border_color="#cbd5e1",
        ).grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 22), pady=7)

    def _file_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: ctk.StringVar,
        action,
        radio_value: str | None = None,
    ) -> None:
        if radio_value:
            label_widget = ctk.CTkRadioButton(
                parent,
                text=label,
                value=radio_value,
                variable=self.source_mode,
                radiobutton_width=16,
                radiobutton_height=16,
                border_width_checked=5,
                font=ctk.CTkFont("Microsoft YaHei UI", 13),
            )
        else:
            label_widget = ctk.CTkLabel(parent, text=label, text_color="#334155", font=ctk.CTkFont("Microsoft YaHei UI", 13))
        label_widget.grid(row=row, column=0, sticky="w", padx=(22, 14), pady=7)
        ctk.CTkEntry(parent, textvariable=variable, height=38, corner_radius=10, border_color="#cbd5e1").grid(
            row=row, column=1, sticky="ew", pady=7
        )
        ctk.CTkButton(
            parent,
            text="选择",
            width=92,
            height=38,
            corner_radius=10,
            fg_color="#f8fafc",
            hover_color="#e2e8f0",
            text_color=self.colors["ink"],
            border_color="#cbd5e1",
            border_width=1,
            command=action,
        ).grid(row=row, column=2, sticky="ew", padx=(10, 22), pady=7)

    def _build_log_card(self, parent: ctk.CTkFrame) -> None:
        card = self._card(parent, 3)
        ctk.CTkLabel(card, text="运行日志", text_color=self.colors["ink"], font=ctk.CTkFont("Microsoft YaHei UI", 18, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", padx=22, pady=(20, 0)
        )
        ctk.CTkLabel(card, textvariable=self.status_text, text_color=self.colors["muted"], font=ctk.CTkFont("Microsoft YaHei UI", 12)).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=22, pady=(4, 10)
        )
        ctk.CTkProgressBar(card, variable=self._progress_var(), height=12, corner_radius=999).grid(
            row=2, column=0, columnspan=3, sticky="ew", padx=22, pady=(0, 12)
        )
        self.log_text = ctk.CTkTextbox(
            card,
            fg_color=self.colors["soft"],
            text_color="#334155",
            border_color=self.colors["line"],
            border_width=1,
            corner_radius=14,
            font=ctk.CTkFont("Microsoft YaHei UI", 12),
            height=170,
        )
        self.log_text.grid(row=3, column=0, columnspan=3, sticky="ew", padx=22, pady=(0, 22))

    def _build_result_card(self, parent: ctk.CTkFrame) -> None:
        card = ctk.CTkFrame(parent, fg_color=self.colors["surface"], corner_radius=22)
        self.result_card = card
        card.pack(fill="both", expand=True)
        ctk.CTkLabel(card, text="结果交付", text_color=self.colors["ink"], font=ctk.CTkFont("Microsoft YaHei UI", 18, "bold")).pack(
            anchor="w", padx=22, pady=(20, 0)
        )
        ctk.CTkLabel(
            card,
            text="普通用户只需要打开这两个结果文件。",
            text_color=self.colors["muted"],
            font=ctk.CTkFont("Microsoft YaHei UI", 12),
        ).pack(anchor="w", padx=22, pady=(4, 16))

        progress = ctk.CTkFrame(card, fg_color=self.colors["soft"], corner_radius=16)
        progress.pack(fill="x", padx=22, pady=(0, 16))
        ctk.CTkLabel(progress, textvariable=self.status_text, text_color="#334155", font=ctk.CTkFont("Microsoft YaHei UI", 13)).pack(
            anchor="w", padx=14, pady=(12, 0)
        )
        ctk.CTkLabel(progress, textvariable=self.progress_percent, text_color=self.colors["muted"], font=ctk.CTkFont("Microsoft YaHei UI", 12)).pack(
            anchor="w", padx=14, pady=(2, 8)
        )
        ctk.CTkProgressBar(progress, variable=self._progress_var(), height=12, corner_radius=999).pack(fill="x", padx=14, pady=(0, 14))

        metrics = ctk.CTkFrame(card, fg_color="transparent")
        metrics.pack(fill="x", padx=22, pady=(0, 18))
        self._metric(metrics, "处理图片", self.total_images).pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._metric(metrics, "问题结果", self.problem_count).pack(side="left", fill="x", expand=True)

        self._result_button(card, "打开主结果表", self._open_main_result, primary=True)
        self._result_button(card, "打开问题结果表", self._open_problem_result, primary=True)
        self._result_button(card, "打开结果文件夹", self._open_run_dir)

        ctk.CTkFrame(card, height=1, fg_color=self.colors["line"]).pack(fill="x", padx=22, pady=18)
        ctk.CTkLabel(card, text="诊断资料", text_color="#334155", font=ctk.CTkFont("Microsoft YaHei UI", 13, "bold")).pack(anchor="w", padx=22)
        ctk.CTkLabel(
            card,
            text="日志、checkpoint、二次审核明细会保存在本次运行文件夹，仅用于排查和反馈修正。",
            text_color=self.colors["muted"],
            wraplength=300,
            justify="left",
            font=ctk.CTkFont("Microsoft YaHei UI", 12),
        ).pack(anchor="w", padx=22, pady=(4, 12))
        self._result_button(card, "打开诊断文件夹", self._open_run_dir)

    def _metric(self, parent: ctk.CTkFrame, label: str, variable: ctk.StringVar) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, fg_color=self.colors["soft"], corner_radius=16)
        ctk.CTkLabel(frame, textvariable=variable, text_color=self.colors["ink"], font=ctk.CTkFont("Segoe UI", 28, "bold")).pack(
            anchor="w", padx=14, pady=(12, 0)
        )
        ctk.CTkLabel(frame, text=label, text_color=self.colors["muted"], font=ctk.CTkFont("Microsoft YaHei UI", 12)).pack(
            anchor="w", padx=14, pady=(0, 12)
        )
        return frame

    def _result_button(self, parent: ctk.CTkFrame, text: str, command, primary: bool = False) -> None:
        ctk.CTkButton(
            parent,
            text=text,
            height=42,
            corner_radius=12,
            fg_color=self.colors["blue"] if primary else "#ffffff",
            hover_color=self.colors["blue_dark"] if primary else "#f1f5f9",
            text_color="#ffffff" if primary else self.colors["ink"],
            border_width=0 if primary else 1,
            border_color=self.colors["line"],
            font=ctk.CTkFont("Microsoft YaHei UI", 14, "bold" if primary else "normal"),
            command=command,
        ).pack(fill="x", padx=22, pady=(0, 10))

    def _progress_var(self) -> ctk.DoubleVar:
        if not hasattr(self, "progress_fraction"):
            self.progress_fraction = ctk.DoubleVar(value=0.0)
        return self.progress_fraction

    def _choose_input_xlsx(self) -> None:
        path = filedialog.askopenfilename(title="选择源数据 Excel", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.input_xlsx.set(path)
            self.source_mode.set("excel")

    def _choose_ids_file(self) -> None:
        path = filedialog.askopenfilename(title="选择商品 ID 文件", filetypes=[("Text/CSV", "*.txt *.csv"), ("All", "*.*")])
        if path:
            self.ids_file.set(path)
            self.source_mode.set("ids")

    def _choose_har_file(self) -> None:
        path = filedialog.askopenfilename(title="选择 HAR 文件", filetypes=[("HAR", "*.har"), ("All", "*.*")])
        if path:
            self.har_file.set(path)

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def _choose_template_xlsx(self) -> None:
        path = filedialog.askopenfilename(title="选择示例模板", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if path:
            self.template_xlsx.set(path)

    def _choose_correction_xlsx(self) -> None:
        path = filedialog.askopenfilename(title="选择修正反馈表", filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if path:
            self.correction_xlsx.set(path)

    def _start_run(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在运行", "当前任务还没有结束。")
            return
        try:
            self._validate_inputs()
        except ValueError as exc:
            messagebox.showerror("无法开始", str(exc))
            return

        self.log_text.delete("1.0", "end")
        self.total_images.set("-")
        self.problem_count.set("-")
        self.current_outputs = None
        self.last_clean_log = ""
        self.status_text.set("正在准备运行")
        self._set_progress(0)
        self.run_button.configure(state="disabled")
        self._emit("log", "已点击开始运行，正在启动任务...")
        self.worker = threading.Thread(target=self._run_pipeline, daemon=True)
        self.worker.start()

    def _validate_inputs(self) -> None:
        if self.source_mode.get() == "excel":
            path = Path(self.input_xlsx.get())
            if not path.exists():
                raise ValueError("请选择已有 Excel 文件。")
        else:
            ids = Path(self.ids_file.get())
            har = Path(self.har_file.get())
            if not ids.exists():
                raise ValueError("请选择商品 ID 文件。")
            if not har.exists():
                raise ValueError("商品 ID 导出需要可用的 HAR 登录文件。")
        if not DETECT_SCRIPT.exists():
            raise ValueError(f"找不到判断脚本：{DETECT_SCRIPT}")
        if self.source_mode.get() == "ids" and not EXPORT_SCRIPT.exists():
            raise ValueError(f"找不到导出脚本：{EXPORT_SCRIPT}")
        batch = self.batch_size.get().strip()
        if not batch.isdigit() or int(batch) < 1:
            raise ValueError("检查点批量必须是正整数。")
        if self.llm_enabled.get():
            if not self.llm_api_url.get().strip():
                raise ValueError("启用大模型辅助时，请填写 API 地址。")
            if not self.llm_api_key.get().strip():
                raise ValueError("启用大模型辅助时，请填写 API Key。")
            if not self.llm_model.get().strip():
                raise ValueError("启用大模型辅助时，请填写模型名称。")

    def _run_pipeline(self) -> None:
        run_id = time.strftime("%Y%m%d_%H%M%S")
        run_dir = Path(self.output_dir.get()) / run_id
        run_dir.mkdir(parents=True, exist_ok=True)
        log_file = run_dir / "运行日志.txt"

        try:
            input_xlsx = Path(self.input_xlsx.get())
            if self.source_mode.get() == "ids":
                self._emit("status", "正在导出源数据")
                self._emit("progress", "5")
                export_dir = run_dir / "source_exports"
                self._run_command(
                    [
                        sys.executable,
                        str(EXPORT_SCRIPT),
                        "--har",
                        self.har_file.get(),
                        "--ids-file",
                        self.ids_file.get(),
                        "--output-dir",
                        str(export_dir),
                    ],
                    log_file,
                )
                exported = sorted(export_dir.glob("*.xlsx"), key=lambda item: item.stat().st_mtime)
                if not exported:
                    raise RuntimeError("源数据导出完成，但没有找到 Excel 文件。")
                input_xlsx = exported[-1]
                self._emit("clean_log", f"使用导出的源数据：{input_xlsx}")

            self._emit("status", "正在判断死图")
            self._emit("progress", "10")
            main_xlsx = run_dir / f"{safe_stem(input_xlsx)}_主结果表.xlsx"
            checkpoint = run_dir / f"{safe_stem(input_xlsx)}.checkpoint.jsonl"
            cmd = [
                sys.executable,
                str(DETECT_SCRIPT),
                str(input_xlsx),
                "--output-dir",
                str(run_dir),
                "--output-xlsx",
                str(main_xlsx),
                "--checkpoint-jsonl",
                str(checkpoint),
                "--batch-size",
                self.batch_size.get().strip(),
            ]
            if self.template_xlsx.get().strip() and Path(self.template_xlsx.get()).exists():
                cmd.extend(["--template-xlsx", self.template_xlsx.get()])
            if self.correction_xlsx.get().strip() and Path(self.correction_xlsx.get()).exists():
                cmd.extend(["--correction-xlsx", self.correction_xlsx.get()])
            if self.llm_enabled.get():
                self._emit("clean_log", f"大模型辅助已启用：{self.llm_model.get().strip()}")

            self._run_command(cmd, log_file)

            self._emit("status", "正在生成问题结果表")
            self._emit("progress", "95")
            problem_xlsx, total_images, problem_count = create_problem_workbook(main_xlsx)
            self.current_outputs = RunOutputs(run_dir, main_xlsx, problem_xlsx, checkpoint, log_file)
            self._emit("metrics", f"{total_images}|{problem_count}")
            self._emit("clean_log", f"主结果表：{main_xlsx}")
            self._emit("clean_log", f"问题结果表：{problem_xlsx}")
            self._emit("status", "运行完成")
            self._emit("progress", "100")
        except Exception as exc:
            self._emit("status", "运行失败")
            self._emit("clean_log", f"错误：{exc}")
            message = str(exc)
            self.after(0, lambda: messagebox.showerror("运行失败", message))
        finally:
            self._emit("done", "")

    def _run_command(self, cmd: list[str], log_file: Path) -> None:
        self._emit("clean_log", "已启动后台任务")
        with log_file.open("a", encoding="utf-8") as log:
            log.write("执行：" + " ".join(quote_arg(part) for part in cmd) + "\n")
            env = os.environ.copy()
            env["PYTHONUTF8"] = "1"
            env["PYTHONIOENCODING"] = "utf-8"
            if self.llm_enabled.get():
                env["DEAD_IMAGE_LLM_ENABLED"] = "1"
                env["DEAD_IMAGE_LLM_API_URL"] = self.llm_api_url.get().strip()
                env["DEAD_IMAGE_LLM_API_KEY"] = self.llm_api_key.get().strip()
                env["DEAD_IMAGE_LLM_MODEL"] = self.llm_model.get().strip()
            else:
                env["DEAD_IMAGE_LLM_ENABLED"] = "0"
            process = subprocess.Popen(
                cmd,
                cwd=ROOT,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
            )
            assert process.stdout is not None
            for line in process.stdout:
                text = repair_mojibake(line.rstrip())
                log.write(text + "\n")
                progress = parse_progress(text)
                if progress:
                    self._emit("status", progress[0])
                    self._emit("progress", str(progress[1]))
                    self._emit("clean_log", progress[0])
            code = process.wait()
            if code != 0:
                raise RuntimeError(f"命令执行失败，退出码 {code}。请查看运行日志。")

    def _emit(self, kind: str, text: str) -> None:
        self.log_queue.put((kind, text))

    def _drain_log_queue(self) -> None:
        try:
            while True:
                kind, text = self.log_queue.get_nowait()
                if kind in {"log", "clean_log"}:
                    if kind == "clean_log" and text == self.last_clean_log:
                        continue
                    if kind == "clean_log":
                        self.last_clean_log = text
                    self.log_text.insert("end", text + "\n")
                    self.log_text.see("end")
                elif kind == "status":
                    self.status_text.set(text)
                elif kind == "metrics":
                    total, problem = text.split("|", 1)
                    self.total_images.set(total)
                    self.problem_count.set(problem)
                elif kind == "progress":
                    try:
                        self._set_progress(float(text))
                    except ValueError:
                        pass
                elif kind == "done":
                    self.run_button.configure(state="normal")
        except queue.Empty:
            pass
        self.after(120, self._drain_log_queue)

    def _set_progress(self, value: float) -> None:
        value = max(0.0, min(100.0, value))
        self.progress_percent.set(f"{value:.1f}%")
        self._progress_var().set(value / 100)

    def _open_main_result(self) -> None:
        if self.current_outputs:
            self._open_path(self.current_outputs.main_xlsx)

    def _open_problem_result(self) -> None:
        if self.current_outputs and self.current_outputs.problem_xlsx:
            self._open_path(self.current_outputs.problem_xlsx)

    def _open_run_dir(self) -> None:
        if self.current_outputs:
            self._open_path(self.current_outputs.run_dir)
        else:
            self._open_path(Path(self.output_dir.get()))

    def _open_path(self, path: Path) -> None:
        try:
            if path.exists():
                os.startfile(path)
            else:
                messagebox.showinfo("未找到", f"路径不存在：{path}")
        except OSError as exc:
            messagebox.showerror("无法打开", str(exc))


def main() -> None:
    app = ModernDeadImageWorkbench()
    app.mainloop()


if __name__ == "__main__":
    main()
