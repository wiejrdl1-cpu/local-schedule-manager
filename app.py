from __future__ import annotations

import calendar
import json
import os
import secrets
import sys
import tkinter as tk
import uuid
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk
from tkinter import font as tkfont
from typing import Callable

import holidays

from database import Database
from excel_import import (
    ImportItem,
    apply_import,
    apply_task_import,
    create_task_template,
    preview_import,
    preview_task_import,
)
from notifications import show_windows_notification
from security import InvalidPasswordError


APP_TITLE = "간편 플래너"
DATE_HINT = "YYYY-MM-DD 또는 YYYY-MM-DD HH:MM"
PALETTE = {
    "background": "#F4FBEF",
    "surface": "#FFFFFF",
    "surface_green": "#ECF8E6",
    "green_soft": "#D8F0CD",
    "green": "#8BCB88",
    "green_dark": "#3F6B46",
    "text": "#294332",
    "muted": "#66806C",
    "border": "#BFDDB8",
}
DEFAULT_STAGES = [
    {"id": "application", "name": "신청", "branches": []},
    {"id": "public-request", "name": "공적요청", "branches": []},
    {"id": "investigation", "name": "조사중", "branches": []},
    {"id": "investigation-complete", "name": "조사완료", "branches": []},
]
DEFAULT_TASK_STAGES = [
    {"id": "task-received", "name": "접수", "branches": []},
    {"id": "task-working", "name": "작업중", "branches": []},
    {"id": "task-complete", "name": "완료", "branches": []},
]
PROGRESS_COLORS = {"completed": "#E34F4F", "current": "#3E82D7", "pending": "#C9CECA"}


def resource_path(relative_path: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / relative_path


def parse_deadline(value: str) -> datetime:
    text = value.strip()
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    raise ValueError(f"날짜는 {DATE_HINT} 형식으로 입력하세요.")


def dday_text(value: str, completed: bool = False) -> str:
    if completed:
        return "완료"
    try:
        target = parse_deadline(value).date()
    except ValueError:
        return "날짜 확인"
    delta = (target - date.today()).days
    if delta > 0:
        return f"D-{delta}"
    if delta == 0:
        return "D-Day"
    return f"{abs(delta)}일 초과"


def urgency_tag(value: str, completed: bool = False) -> str:
    if completed:
        return "completed"
    try:
        delta = (parse_deadline(value).date() - date.today()).days
    except ValueError:
        return "normal"
    if delta < 0:
        return "overdue"
    if delta == 0:
        return "today"
    if delta <= 3:
        return "soon3"
    if delta <= 7:
        return "soon7"
    return "normal"


class FormDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        title: str,
        fields: list[tuple[str, str, str]],
        on_save: Callable[[dict[str, str]], None],
    ):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(True, True)
        self.minsize(520, 360)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.entries: dict[str, tk.Widget] = {}
        self.on_save = on_save
        body = ttk.Frame(self, padding=18)
        body.grid(sticky="nsew")
        body.columnconfigure(1, weight=1)
        for row, (key, label, default) in enumerate(fields):
            ttk.Label(body, text=label).grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=6)
            if key in {"memo", "content", "reason", "details"}:
                widget = tk.Text(body, width=46, height=4, wrap="word", font=getattr(parent, "input_font", None))
                widget.insert("1.0", default)
            else:
                widget = ttk.Entry(body, width=46)
                widget.insert(0, default)
            is_multiline = isinstance(widget, tk.Text)
            widget.grid(row=row, column=1, sticky="nsew" if is_multiline else "ew", pady=6)
            if is_multiline:
                body.rowconfigure(row, weight=1)
            self.entries[key] = widget
        buttons = ttk.Frame(body)
        buttons.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="취소", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="저장", style="Accent.TButton", command=self._save).pack(side="right")
        self.bind("<Escape>", lambda _e: self.destroy())
        self.grab_set()
        self.after(50, lambda: next(iter(self.entries.values())).focus_set())

    def _save(self) -> None:
        data: dict[str, str] = {}
        for key, widget in self.entries.items():
            if isinstance(widget, tk.Text):
                data[key] = widget.get("1.0", "end").strip()
            else:
                data[key] = widget.get().strip()
        try:
            self.on_save(data)
        except Exception as exc:
            messagebox.showerror("저장할 수 없음", str(exc), parent=self)
            return
        self.destroy()


class ScheduleApp(tk.Tk):
    def __init__(self, db: Database):
        super().__init__()
        self.db = db
        self.import_path: Path | None = None
        self.import_items: list[ImportItem] = []
        self.stage_config = self._load_stage_config()
        self.task_stage_config = self._load_task_stage_config()
        today = date.today()
        self.calendar_month = today.replace(day=1)
        self.calendar_selected_date = today
        self._progress_canvases: list[tk.Canvas] = []
        self._complaint_progress_buttons: list[ttk.Button] = []
        self._complaint_progress: dict[str, tuple[dict[str, object], str]] = {}
        self._task_progress_canvases: list[tk.Canvas] = []
        self._task_progress_buttons: list[ttk.Button] = []
        self._task_progress: dict[str, tuple[dict[str, object], str]] = {}
        self.title(APP_TITLE)
        icon_path = resource_path("assets/app_icon.ico")
        if icon_path.exists():
            try:
                self.iconbitmap(default=str(icon_path))
            except tk.TclError:
                pass
        self.geometry("1500x900")
        self.minsize(1200, 760)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._configure_style()
        self._build_ui()
        self.refresh_all()
        self.after(900, self._notify_due_items)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        installed_fonts = set(tkfont.families(self))
        self.ui_font = next(
            (name for name in ("NanumGothic", "나눔고딕", "Arial Rounded MT Bold", "맑은 고딕") if name in installed_fonts),
            "맑은 고딕",
        )
        self.input_font = tkfont.Font(self, family=self.ui_font, size=10, weight="normal")
        for named_font in ("TkDefaultFont", "TkTextFont", "TkFixedFont", "TkMenuFont"):
            try:
                tkfont.nametofont(named_font, root=self).configure(
                    family=self.ui_font,
                    size=10,
                    weight="normal",
                )
            except tk.TclError:
                pass
        self.option_add("*Font", self.input_font)
        self.option_add("*Entry.font", self.input_font)
        self.option_add("*Text.font", self.input_font)
        self.option_add("*Text.borderWidth", 1)
        self.option_add("*Text.relief", "solid")
        self.option_add("*Text.highlightThickness", 1)
        self.option_add("*Text.highlightBackground", "#879987")
        self.option_add("*Text.highlightColor", PALETTE["green_dark"])
        self.option_add("*TCombobox*Listbox.font", self.input_font)
        self.configure(background=PALETTE["background"])
        style.configure("TFrame", background=PALETTE["background"])
        style.configure("TLabel", font=(self.ui_font, 11, "bold"), background=PALETTE["background"], foreground=PALETTE["text"])
        style.configure(
            "TEntry",
            font=self.input_font,
            fieldbackground="#FFFFFF",
            borderwidth=1,
            relief="solid",
            bordercolor="#879987",
            lightcolor="#879987",
            darkcolor="#879987",
        )
        style.map("TEntry", bordercolor=[("focus", PALETTE["green_dark"])])
        style.configure(
            "TCombobox",
            font=self.input_font,
            fieldbackground="#FFFFFF",
            borderwidth=1,
            relief="solid",
            bordercolor="#879987",
            lightcolor="#879987",
            darkcolor="#879987",
        )
        style.map("TCombobox", bordercolor=[("focus", PALETTE["green_dark"])])
        style.configure("TCheckbutton", font=(self.ui_font, 10, "bold"), background=PALETTE["background"], foreground=PALETTE["text"])
        style.configure("TRadiobutton", font=(self.ui_font, 10, "bold"), background=PALETTE["background"], foreground=PALETTE["text"])
        style.configure("TLabelframe", background=PALETTE["surface"], bordercolor=PALETTE["border"])
        style.configure("TLabelframe.Label", font=(self.ui_font, 10, "bold"), background=PALETTE["background"], foreground=PALETTE["green_dark"])
        style.configure("Title.TLabel", font=(self.ui_font, 25, "bold"), foreground=PALETTE["green_dark"])
        style.configure("Subtitle.TLabel", font=(self.ui_font, 12, "bold"), foreground="#465B4B")
        style.configure("Card.TFrame", background="#FFFFFF", relief="solid", borderwidth=1, bordercolor="#C8D0C5")
        style.configure("CardValueBlack.TLabel", font=(self.ui_font, 32, "bold"), foreground="#202020", background="#FFFFFF")
        style.configure("CardValueRed.TLabel", font=(self.ui_font, 32, "bold"), foreground="#D74343", background="#FFFFFF")
        style.configure("CardValueOrange.TLabel", font=(self.ui_font, 32, "bold"), foreground="#E47D26", background="#FFFFFF")
        style.configure("CardValueBlue.TLabel", font=(self.ui_font, 32, "bold"), foreground="#3478C5", background="#FFFFFF")
        style.configure("CardLabel.TLabel", font=(self.ui_font, 16, "bold"), foreground="#303830", background="#FFFFFF")
        style.configure("TButton", font=(self.ui_font, 10, "bold"), padding=(10, 6), background=PALETTE["background"], foreground="#111111", bordercolor=PALETTE["border"])
        style.map("TButton", background=[("active", PALETTE["background"]), ("pressed", PALETTE["background"])], foreground=[("active", "#111111")])
        style.configure("Accent.TButton", font=(self.ui_font, 9, "bold"), background=PALETTE["background"], foreground="#111111", bordercolor=PALETTE["border"])
        style.map("Accent.TButton", background=[("active", PALETTE["background"]), ("pressed", PALETTE["background"])], foreground=[("active", "#111111")])
        style.configure("ExcelAction.TButton", font=(self.ui_font, 10, "bold"), padding=(10, 6))
        style.configure("Edit.TButton", font=(self.ui_font, 10, "bold"), background=PALETTE["background"], foreground="#111111", bordercolor=PALETTE["border"])
        style.map("Edit.TButton", background=[("active", PALETTE["background"]), ("pressed", PALETTE["background"])], foreground=[("active", "#111111")])
        style.configure("SelectedEdit.TButton", font=(self.ui_font, 10, "bold"), background="#A9DDA4", foreground="#173B23", bordercolor="#71906F")
        style.map("SelectedEdit.TButton", background=[("active", "#A9DDA4"), ("pressed", "#8CCB87")], foreground=[("active", "#173B23")])
        style.configure("Nav.TButton", font=(self.ui_font, 10, "bold"), padding=(14, 10), background=PALETTE["background"], foreground="#111111", bordercolor=PALETTE["border"])
        style.map("Nav.TButton", background=[("active", PALETTE["background"]), ("pressed", PALETTE["background"])], foreground=[("active", "#111111")])
        style.configure("NavActive.TButton", font=(self.ui_font, 10, "bold"), padding=(14, 10), background=PALETTE["green_dark"], foreground="#FFFFFF", bordercolor=PALETTE["green_dark"])
        style.map("NavActive.TButton", background=[("active", PALETTE["green_dark"]), ("pressed", PALETTE["green_dark"])], foreground=[("active", "#FFFFFF")])
        style.layout("Hidden.TNotebook.Tab", [])
        style.configure("Hidden.TNotebook", background=PALETTE["background"], borderwidth=0, tabmargins=0)
        style.configure("Treeview", rowheight=100, font=(self.ui_font, 9), background=PALETTE["surface"], fieldbackground=PALETTE["surface"], foreground=PALETTE["text"], bordercolor="#B8C9B5", borderwidth=1, relief="solid")
        style.map("Treeview", background=[("selected", "#A9DDA4")], foreground=[("selected", "#173B23")])
        style.configure("Treeview.Heading", font=(self.ui_font, 9, "bold"), background=PALETTE["green_soft"], foreground=PALETTE["green_dark"], bordercolor="#D9E5D6")
        style.map("Treeview.Heading", background=[("active", "#C4E7B9")])
        style.configure("Calendar.Treeview", rowheight=34, font=(self.ui_font, 9), background="#FFFFFF", fieldbackground="#FFFFFF", foreground=PALETTE["text"], bordercolor="#B8C9B5", borderwidth=1)
        style.map("Calendar.Treeview", background=[("selected", "#A9DDA4")], foreground=[("selected", "#173B23")])
        style.configure("Calendar.Treeview.Heading", font=(self.ui_font, 9, "bold"), background=PALETTE["green_soft"], foreground=PALETTE["green_dark"])

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=(22, 18))
        root.pack(fill="both", expand=True)
        header = ttk.Frame(root)
        header.pack(fill="x")
        ttk.Label(header, text=APP_TITLE, style="Title.TLabel").pack(side="left")
        ttk.Label(header, text="© 2026 이주삼", style="Subtitle.TLabel").pack(side="right", anchor="ne")

        self.summary_vars = {
            key: tk.StringVar(value="0")
            for key in ("민원", "수시업무", "오늘의 일정", "오늘 마감", "3일 이내", "기한 초과")
        }
        cards = ttk.Frame(root)
        cards.pack(fill="x", pady=(18, 16))
        card_value_colors = {
            "민원": "#202020",
            "수시업무": "#202020",
            "오늘의 일정": "#202020",
            "오늘 마감": "#D74343",
            "3일 이내": "#E47D26",
            "기한 초과": "#3478C5",
        }
        for col, (label, variable) in enumerate(self.summary_vars.items()):
            card = ttk.Frame(cards, padding=(20, 16), style="Card.TFrame")
            last_col = len(self.summary_vars) - 1
            card.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 5, 0 if col == last_col else 5))
            cards.columnconfigure(col, weight=1)
            value_label = tk.Label(
                card,
                textvariable=variable,
                font=(self.ui_font, 32, "bold"),
                foreground=card_value_colors[label],
                background="#FFFFFF",
                anchor="center",
                cursor="hand2",
            )
            value_label.pack(fill="x")
            name_label = tk.Label(
                card,
                text=label,
                font=(self.ui_font, 16, "bold"),
                foreground="#202020",
                background="#FFFFFF",
                anchor="center",
                cursor="hand2",
            )
            name_label.pack(fill="x")
            destination = 1 if label == "민원" else 2 if label == "수시업무" else 3 if label == "오늘의 일정" else 0
            for clickable in (card, value_label, name_label):
                clickable.bind("<Button-1>", lambda _event, page=destination: self.show_page(page))

        navigation = ttk.Frame(root)
        navigation.pack(fill="x", pady=(0, 14))

        self.notebook = ttk.Notebook(root, style="Hidden.TNotebook")
        self.notebook.pack(fill="both", expand=True)
        self.today_tab = ttk.Frame(self.notebook, padding=12)
        self.complaints_tab = ttk.Frame(self.notebook, padding=12)
        self.tasks_tab = ttk.Frame(self.notebook, padding=12)
        self.calendar_tab = ttk.Frame(self.notebook, padding=12)
        self.completed_tab = ttk.Frame(self.notebook, padding=12)
        self.trash_tab = ttk.Frame(self.notebook, padding=12)
        pages = (
            (self.today_tab, "오늘의 업무"),
            (self.complaints_tab, "새올 민원 처리"),
            (self.tasks_tab, "수시업무 처리"),
            (self.calendar_tab, "월간 일정"),
            (self.completed_tab, "완료 내역"),
            (self.trash_tab, "휴지통"),
        )
        self.nav_buttons: list[ttk.Button] = []
        for index, (tab, text) in enumerate(pages):
            self.notebook.add(tab, text=text)
            navigation.columnconfigure(index, weight=1)
            button = ttk.Button(
                navigation,
                text=text,
                style="Nav.TButton",
                command=lambda page=index: self.show_page(page),
            )
            button.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 4, 0 if index == len(pages) - 1 else 4))
            self.nav_buttons.append(button)
        self._build_today_tab()
        self._build_complaints_tab()
        self._build_tasks_tab()
        self._build_calendar_tab()
        self._build_completed_tab()
        self._build_trash_tab()
        self.show_page(0)

    def show_page(self, index: int) -> None:
        self.notebook.select(index)
        for button_index, button in enumerate(self.nav_buttons):
            button.configure(style="NavActive.TButton" if button_index == index else "Nav.TButton")
        self.after_idle(self._draw_complaint_progress)
        self.after_idle(self._draw_task_progress)

    def _load_stage_config(self) -> list[dict[str, object]]:
        raw = self.db.get_setting("processing_stage_config")
        value: list[dict[str, object]] | None = None
        if raw:
            try:
                loaded = json.loads(raw)
                if isinstance(loaded, list) and loaded:
                    value = loaded
            except (json.JSONDecodeError, TypeError):
                pass
        if value is not None and [str(stage.get("name", "")) for stage in value] == ["공적자료요청", "공적회신", "조사중", "조사완료"]:
            value = None
        if value is None:
            value = json.loads(json.dumps(DEFAULT_STAGES, ensure_ascii=False))
        changed = False
        for stage in value:
            if not stage.get("id"):
                stage["id"] = uuid.uuid4().hex
                changed = True
            normalized_branches = []
            for branch in stage.get("branches", []):
                if isinstance(branch, str):
                    normalized_branches.append({"id": uuid.uuid4().hex, "name": branch})
                    changed = True
                else:
                    if not branch.get("id"):
                        branch["id"] = uuid.uuid4().hex
                        changed = True
                    normalized_branches.append(branch)
            stage["branches"] = normalized_branches
        if changed:
            self.db.set_setting("processing_stage_config", json.dumps(value, ensure_ascii=False))
        return value

    def _save_stage_config(self) -> None:
        self.db.set_setting(
            "processing_stage_config",
            json.dumps(self.stage_config, ensure_ascii=False),
        )

    def _load_task_stage_config(self) -> list[dict[str, object]]:
        raw = self.db.get_setting("task_stage_config")
        if raw:
            try:
                loaded = json.loads(raw)
                if isinstance(loaded, list) and loaded:
                    changed = False
                    for stage in loaded:
                        if not stage.get("id"):
                            stage["id"] = uuid.uuid4().hex
                            changed = True
                        normalized_branches = []
                        for branch in stage.get("branches", []):
                            if isinstance(branch, str):
                                normalized_branches.append({"id": uuid.uuid4().hex, "name": branch})
                                changed = True
                            else:
                                if not branch.get("id"):
                                    branch["id"] = uuid.uuid4().hex
                                    changed = True
                                normalized_branches.append(branch)
                        if stage.get("branches") != normalized_branches:
                            changed = True
                        stage["branches"] = normalized_branches
                    if changed:
                        self.db.set_setting("task_stage_config", json.dumps(loaded, ensure_ascii=False))
                    return loaded
            except (json.JSONDecodeError, TypeError):
                pass
        return json.loads(json.dumps(DEFAULT_TASK_STAGES, ensure_ascii=False))

    def _save_task_stage_config(self) -> None:
        self.db.set_setting("task_stage_config", json.dumps(self.task_stage_config, ensure_ascii=False))

    @staticmethod
    def _item_stage_config(raw: object, fallback: list[dict[str, object]]) -> list[dict[str, object]]:
        if raw:
            try:
                loaded = json.loads(str(raw))
                if isinstance(loaded, list) and loaded:
                    return loaded
            except (json.JSONDecodeError, TypeError):
                pass
        return fallback

    def _complaint_stages(self, complaint: dict[str, object]) -> list[dict[str, object]]:
        return self._item_stage_config(complaint.get("stage_config"), self.stage_config)

    def _task_stages(self, task: dict[str, object]) -> list[dict[str, object]]:
        return self._item_stage_config(task.get("stage_config"), self.task_stage_config)

    def _task_progress_state(self, task: dict[str, object]) -> dict[str, object]:
        stages = self._task_stages(task)
        try:
            state = json.loads(str(task.get("progress_state") or "{}"))
        except json.JSONDecodeError:
            state = {}
        valid = set(self._progress_node_keys(stages))
        if state.get("current") in valid:
            completed = [key for key in state.get("completed", []) if key in valid]
            return {"current": state["current"], "completed": completed}
        stage_text = str(task.get("processing_stage", ""))
        index = next((i for i, stage in enumerate(stages) if str(stage["name"]) in stage_text), 0)
        stage = stages[index]
        branches = stage.get("branches", [])
        current = f"b:{stage['id']}:{branches[0]['id']}" if branches else f"s:{stage['id']}"
        for branch in branches:
            if str(branch["name"]) in stage_text:
                current = f"b:{stage['id']}:{branch['id']}"
                break
        previous = self._progress_node_keys(stages[:index])
        return {
            "current": current,
            "completed": previous,
        }

    def _task_stage_text(self, key: str, stages: list[dict[str, object]] | None = None) -> str:
        stages = stages or self.task_stage_config
        return self._definition_text_from_key(key, stages)

    def _toggle_task_progress(self, task_id: int, node_key: str, is_completed: bool) -> None:
        task = self.db.task_by_id(task_id)
        if not task:
            return
        stages = self._task_stages(task)
        state = self._task_progress_state(task)
        nodes = self._progress_node_keys(stages)
        if node_key not in nodes:
            return
        completed = set(state.get("completed", []))
        selected_index = nodes.index(node_key)
        if is_completed:
            completed.update(nodes[: selected_index + 1])
            pending = [key for key in nodes if key not in completed]
            current = pending[0] if pending else node_key
        else:
            completed.difference_update(nodes[selected_index:])
            current = node_key
        self.db.update_task_progress_state(
            task_id,
            {"current": current, "completed": sorted(completed)},
            self._task_stage_text(current, stages),
        )
        self.refresh_all()

    def _stage_label(self, index: int) -> str:
        return f"{index + 1}. {self.stage_config[index]['name']}"

    def _stage_index(self, current_stage: str, stages: list[dict[str, object]] | None = None) -> int:
        stages = stages or self.stage_config
        for index, stage in enumerate(stages):
            if str(stage["name"]) in current_stage:
                return index
        aliases = {
            "접수": "신청",
            "공적자료요청": "공적요청",
            "공적회신": "공적요청",
        }
        for old_name, new_name in aliases.items():
            if old_name in current_stage:
                return next(
                    (index for index, stage in enumerate(stages) if stage["name"] == new_name),
                    -1,
                )
        return -1

    def _progress_state(self, complaint: dict[str, object]) -> dict[str, object]:
        stages = self._complaint_stages(complaint)
        try:
            state = json.loads(str(complaint.get("progress_state") or "{}"))
        except json.JSONDecodeError:
            state = {}
        valid_keys = set()
        for stage in stages:
            valid_keys.add(f"s:{stage['id']}")
            valid_keys.update(f"b:{stage['id']}:{branch['id']}" for branch in stage.get("branches", []))
        if state.get("current") in valid_keys:
            state["completed"] = list(state.get("completed", []))
            return state
        current_stage_text = str(complaint.get("processing_stage", ""))
        stage_index = self._stage_index(current_stage_text, stages)
        if stage_index < 0:
            stage_index = 0
        stage = stages[stage_index]
        current_key = f"s:{stage['id']}"
        for branch in stage.get("branches", []):
            if str(branch["name"]) in current_stage_text:
                current_key = f"b:{stage['id']}:{branch['id']}"
                break
        completed = [f"s:{item['id']}" for item in stages[:stage_index]]
        return {"current": current_key, "completed": completed}

    def _definition_text_from_key(self, key: str, stages: list[dict[str, object]] | None = None) -> str:
        stages = stages or self.stage_config
        parts = key.split(":")
        if len(parts) < 2:
            return str(stages[0]["name"])
        stage = next((item for item in stages if str(item["id"]) == parts[1]), stages[0])
        text = str(stage["name"])
        if parts[0] == "b" and len(parts) == 3:
            branch = next((item for item in stage.get("branches", []) if str(item["id"]) == parts[2]), None)
            if branch:
                text += " > " + str(branch["name"])
        return text

    def _progress_node_keys(self, stages: list[dict[str, object]] | None = None) -> list[str]:
        if stages is None:
            stages = self.stage_config
        keys = []
        for stage in stages:
            branches = stage.get("branches", [])
            if branches:
                keys.extend(f"b:{stage['id']}:{branch['id']}" for branch in branches)
            else:
                keys.append(f"s:{stage['id']}")
        return keys

    def _toggle_progress_node(self, complaint_id: int, node_key: str, is_completed: bool) -> None:
        complaint = self.db.complaint_by_id(complaint_id)
        if not complaint:
            return
        stages = self._complaint_stages(complaint)
        state = self._progress_state(complaint)
        ordered_nodes = self._progress_node_keys(stages)
        if node_key not in ordered_nodes:
            return
        selected_index = ordered_nodes.index(node_key)
        completed = set(state.get("completed", []))
        if is_completed:
            completed.update(ordered_nodes[: selected_index + 1])
            pending = [key for key in ordered_nodes if key not in completed]
            current_key = pending[0] if pending else node_key
        else:
            completed.difference_update(ordered_nodes[selected_index:])
            current_key = node_key
        self.db.update_progress_state(
            complaint_id,
            {"current": current_key, "completed": sorted(completed)},
            self._definition_text_from_key(current_key, stages),
        )
        self.refresh_all()

    def _draw_complaint_progress(self) -> None:
        if not hasattr(self, "complaint_tree") or not self.complaint_tree.winfo_exists():
            return
        for canvas in self._progress_canvases:
            canvas.destroy()
        self._progress_canvases.clear()
        for button in self._complaint_progress_buttons:
            button.destroy()
        self._complaint_progress_buttons.clear()
        row_backgrounds = {
            "normal": "#F5FCEF", "soon7": "#FFF7CF", "soon3": "#FFE4BD",
            "today": "#FFD0D0", "overdue": "#FFBCBC", "completed": "#E8F2E5",
        }
        selected_rows = set(self.complaint_tree.selection())
        for iid, (complaint, tag) in self._complaint_progress.items():
            bounds = self.complaint_tree.bbox(iid, "progress")
            if not bounds:
                continue
            x, y, width, height = bounds
            background = "#A9DDA4" if iid in selected_rows else row_backgrounds.get(tag, "#FFFFFF")
            canvas = tk.Canvas(
                self.complaint_tree,
                width=max(20, width - 4),
                height=max(20, height - 2),
                background=background,
                highlightthickness=0,
            )
            canvas.place(x=x + 2, y=y + 1)
            self._progress_canvases.append(canvas)
            state = self._progress_state(complaint)
            stages = self._complaint_stages(complaint)
            current_key = str(state.get("current", ""))
            completed = set(state.get("completed", []))
            stage_count = max(1, len(stages))
            gap = 5
            usable_width = max(80, width - 16)
            stage_width = max(35, (usable_width - gap * (stage_count - 1)) // stage_count)
            total_width = stage_width * stage_count + gap * (stage_count - 1)
            start_x = max(4, (width - total_width) // 2)
            for index, stage in enumerate(stages):
                stage_key = f"s:{stage['id']}"
                branch_current = current_key.startswith(f"b:{stage['id']}:")
                branches = stage.get("branches", [])
                branch_keys = [f"b:{stage['id']}:{branch['id']}" for branch in branches]
                stage_completed = stage_key in completed or bool(branch_keys) and all(key in completed for key in branch_keys)
                status = "completed" if stage_completed else "current" if current_key == stage_key or branch_current else "pending"
                left = start_x + index * (stage_width + gap)
                canvas.create_rectangle(left, 4, left + stage_width, 25, fill=PROGRESS_COLORS[status], outline="")
                canvas.create_text(left + stage_width / 2, 14, text=str(stage["name"]), fill="#FFFFFF" if status != "pending" else "#4D554F", font=(self.ui_font, 8, "bold"))
                if branches:
                    canvas.create_line(left + stage_width / 2, 25, left + stage_width / 2, 30, fill="#879188")
                    branch_width = max(8, (stage_width - 2 * (len(branches) - 1)) // len(branches))
                    for branch_index, branch in enumerate(branches):
                        branch_key = f"b:{stage['id']}:{branch['id']}"
                        branch_status = "completed" if branch_key in completed else "current" if current_key == branch_key else "pending"
                        branch_left = left + branch_index * (branch_width + 2)
                        canvas.create_rectangle(branch_left, 30, branch_left + branch_width, 45, fill=PROGRESS_COLORS[branch_status], outline="")
                        if branch_width >= 35:
                            canvas.create_text(branch_left + branch_width / 2, 37, text=str(branch["name"]), fill="#FFFFFF" if branch_status != "pending" else "#4D554F", font=(self.ui_font, 7, "bold"))
                        checked = tk.BooleanVar(value=branch_key in completed)
                        check = tk.Checkbutton(
                            canvas,
                            variable=checked,
                            background=background,
                            activebackground=background,
                            selectcolor="#FFFFFF",
                            command=lambda complaint_id=int(complaint["id"]), key=branch_key, variable=checked: self._toggle_progress_node(complaint_id, key, variable.get()),
                        )
                        canvas.create_window(branch_left + branch_width / 2, 57, window=check)
                else:
                    checked = tk.BooleanVar(value=stage_key in completed)
                    check = tk.Checkbutton(
                        canvas,
                        variable=checked,
                        background=background,
                        activebackground=background,
                        selectcolor="#FFFFFF",
                        command=lambda complaint_id=int(complaint["id"]), key=stage_key, variable=checked: self._toggle_progress_node(complaint_id, key, variable.get()),
                    )
                    canvas.create_window(left + stage_width / 2, 51, window=check)
            def select_row(_event: tk.Event, row_id: str = iid) -> None:
                self._select_tree_row(self.complaint_tree, row_id, _event)

            canvas.bind("<Button-1>", select_row)
            canvas.bind("<Double-1>", lambda _event, row_id=iid: (self.complaint_tree.selection_set(row_id), self.edit_complaint_dialog()))
            edit_bounds = self.complaint_tree.bbox(iid, "progress_edit")
            if edit_bounds:
                edit_x, edit_y, edit_width, edit_height = edit_bounds
                edit_button = ttk.Button(
                    self.complaint_tree,
                    text="수정",
                    style="SelectedEdit.TButton" if iid in selected_rows else "Edit.TButton",
                    command=lambda complaint_id=int(complaint["id"]): self.open_item_stage_editor("complaint", complaint_id),
                )
                edit_button.place(x=edit_x + 10, y=edit_y + max(8, (edit_height - 38) // 2), width=max(70, edit_width - 20), height=38)
                self._complaint_progress_buttons.append(edit_button)
        self._draw_tree_grid(self.complaint_tree)

    def _draw_task_progress(self) -> None:
        if not hasattr(self, "task_tree") or not self.task_tree.winfo_exists():
            return
        for canvas in self._task_progress_canvases:
            canvas.destroy()
        self._task_progress_canvases.clear()
        for button in self._task_progress_buttons:
            button.destroy()
        self._task_progress_buttons.clear()
        row_backgrounds = {
            "normal": "#F5FCEF", "soon7": "#FFF7CF", "soon3": "#FFE4BD",
            "today": "#FFD0D0", "overdue": "#FFBCBC", "completed": "#E8F2E5",
        }
        selected_rows = set(self.task_tree.selection())
        for iid, (task, tag) in self._task_progress.items():
            bounds = self.task_tree.bbox(iid, "progress")
            if not bounds:
                continue
            x, y, width, height = bounds
            background = "#A9DDA4" if iid in selected_rows else row_backgrounds.get(tag, "#FFFFFF")
            canvas = tk.Canvas(self.task_tree, width=max(20, width - 4), height=max(20, height - 2), background=background, highlightthickness=0)
            canvas.place(x=x + 2, y=y + 1)
            self._task_progress_canvases.append(canvas)
            state = self._task_progress_state(task)
            stages = self._task_stages(task)
            current = str(state.get("current", ""))
            completed = set(state.get("completed", []))
            count = max(1, len(stages))
            gap = 6
            stage_width = max(45, (max(100, width - 16) - gap * (count - 1)) // count)
            total_width = stage_width * count + gap * (count - 1)
            start_x = max(4, (width - total_width) // 2)
            for index, stage in enumerate(stages):
                key = f"s:{stage['id']}"
                branches = stage.get("branches", [])
                branch_keys = [f"b:{stage['id']}:{branch['id']}" for branch in branches]
                branch_current = current.startswith(f"b:{stage['id']}:")
                stage_completed = key in completed or bool(branch_keys) and all(branch_key in completed for branch_key in branch_keys)
                status = "completed" if stage_completed else "current" if key == current or branch_current else "pending"
                left = start_x + index * (stage_width + gap)
                canvas.create_rectangle(left, 4, left + stage_width, 25, fill=PROGRESS_COLORS[status], outline="")
                canvas.create_text(left + stage_width / 2, 14, text=str(stage["name"]), fill="#FFFFFF" if status != "pending" else "#4D554F", font=(self.ui_font, 8, "bold"))
                if branches:
                    canvas.create_line(left + stage_width / 2, 25, left + stage_width / 2, 30, fill="#879188")
                    branch_width = max(8, (stage_width - 2 * (len(branches) - 1)) // len(branches))
                    for branch_index, branch in enumerate(branches):
                        branch_key = f"b:{stage['id']}:{branch['id']}"
                        branch_status = "completed" if branch_key in completed else "current" if current == branch_key else "pending"
                        branch_left = left + branch_index * (branch_width + 2)
                        canvas.create_rectangle(branch_left, 30, branch_left + branch_width, 45, fill=PROGRESS_COLORS[branch_status], outline="")
                        if branch_width >= 35:
                            canvas.create_text(branch_left + branch_width / 2, 37, text=str(branch["name"]), fill="#FFFFFF" if branch_status != "pending" else "#4D554F", font=(self.ui_font, 7, "bold"))
                        checked = tk.BooleanVar(value=branch_key in completed)
                        check = tk.Checkbutton(
                            canvas,
                            variable=checked,
                            background=background,
                            activebackground=background,
                            selectcolor="#FFFFFF",
                            command=lambda task_id=int(task["id"]), node=branch_key, variable=checked: self._toggle_task_progress(task_id, node, variable.get()),
                        )
                        canvas.create_window(branch_left + branch_width / 2, 57, window=check)
                else:
                    checked = tk.BooleanVar(value=key in completed)
                    check = tk.Checkbutton(
                        canvas,
                        variable=checked,
                        background=background,
                        activebackground=background,
                        selectcolor="#FFFFFF",
                        command=lambda task_id=int(task["id"]), node=key, variable=checked: self._toggle_task_progress(task_id, node, variable.get()),
                    )
                    canvas.create_window(left + stage_width / 2, 51, window=check)
            canvas.bind("<Button-1>", lambda event, row_id=iid: self._select_tree_row(self.task_tree, row_id, event))
            canvas.bind("<Double-1>", lambda _e, row_id=iid: (self.task_tree.selection_set(row_id), self.edit_task_dialog()))
            edit_bounds = self.task_tree.bbox(iid, "progress_edit")
            if edit_bounds:
                edit_x, edit_y, edit_width, edit_height = edit_bounds
                edit_button = ttk.Button(
                    self.task_tree,
                    text="수정",
                    style="SelectedEdit.TButton" if iid in selected_rows else "Edit.TButton",
                    command=lambda task_id=int(task["id"]): self.open_item_stage_editor("task", task_id),
                )
                edit_button.place(x=edit_x + 10, y=edit_y + max(8, (edit_height - 38) // 2), width=max(70, edit_width - 20), height=38)
                self._task_progress_buttons.append(edit_button)
        self._draw_tree_grid(self.task_tree)

    def _draw_tree_grid(self, tree: ttk.Treeview) -> None:
        if not tree.winfo_exists():
            return
        for line in getattr(tree, "_grid_lines", []):
            if line.winfo_exists():
                line.destroy()
        tree._grid_lines = []  # type: ignore[attr-defined]

        configured = tree.cget("displaycolumns")
        if configured in ("#all", ("#all",)):
            display_columns = list(tree.cget("columns"))
        elif isinstance(configured, str):
            display_columns = list(tree.tk.splitlist(configured))
        else:
            display_columns = list(configured)
        if not display_columns:
            return

        visible_rows: list[tuple[str, tuple[int, int, int, int]]] = []
        first_column = display_columns[0]
        for iid in tree.get_children(""):
            bounds = tree.bbox(iid, first_column)
            if bounds:
                visible_rows.append((iid, bounds))

        tree_width = tree.winfo_width()
        tree_height = tree.winfo_height()
        if tree_width <= 1 or tree_height <= 1:
            return

        vertical_positions = {0}
        if visible_rows:
            sample_iid = visible_rows[0][0]
            for column in display_columns:
                bounds = tree.bbox(sample_iid, column)
                if bounds:
                    x, _y, width, _height = bounds
                    vertical_positions.update((x, x + width))
        else:
            x = 0
            for column in display_columns:
                x += int(tree.column(column, "width"))
                vertical_positions.add(x)

        for x in sorted(vertical_positions):
            if 0 <= x < tree_width:
                line = tk.Frame(tree, background="#D9E5D6", width=1, takefocus=0)
                line.place(x=x, y=0, width=1, height=tree_height)
                line.lift()
                tree._grid_lines.append(line)  # type: ignore[attr-defined]

        horizontal_positions: set[int] = set()
        for _iid, (_x, y, _width, height) in visible_rows:
            horizontal_positions.update((y, y + height))
        heading_bottom = max(
            (y + 1 for y in range(min(80, tree_height)) if tree.identify_region(2, y) == "heading"),
            default=0,
        )
        if heading_bottom:
            horizontal_positions.add(heading_bottom)
        for y in sorted(horizontal_positions):
            if 0 <= y < tree_height:
                is_heading_border = y == heading_bottom
                thickness = 2 if is_heading_border else 1
                color = "#71906F" if is_heading_border else "#B8C9B5"
                line = tk.Frame(tree, background=color, height=thickness, takefocus=0)
                line.place(x=0, y=y, width=tree_width, height=thickness)
                line.lift()
                tree._grid_lines.append(line)  # type: ignore[attr-defined]

    def _new_tree(self, parent: tk.Misc, columns: tuple[str, ...], headings: tuple[str, ...], widths: tuple[int, ...]) -> ttk.Treeview:
        wrapper = ttk.Frame(parent)
        wrapper.pack(fill="both", expand=True)
        tree = ttk.Treeview(wrapper, columns=columns, show="headings", selectmode="extended")
        tree._grid_lines = []  # type: ignore[attr-defined]
        def scroll_command(*args: str) -> None:
            tree.yview(*args)
            self.after_idle(lambda: self._redraw_progress_if_needed(tree))

        scrollbar = ttk.Scrollbar(wrapper, orient="vertical", command=scroll_command)

        def yscroll(first: str, last: str) -> None:
            scrollbar.set(first, last)
            self.after_idle(lambda: self._redraw_progress_if_needed(tree))

        tree.configure(yscrollcommand=yscroll)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        for column, heading, width in zip(columns, headings, widths):
            tree.heading(column, text=heading, command=lambda selected=column: self._sort_tree(tree, selected))
            tree.column(column, width=width, minwidth=60, anchor="center")
        tree.tag_configure("normal", background="#F5FCEF")
        tree.tag_configure("soon7", background="#fff7cf")
        tree.tag_configure("soon3", background="#ffe4bd")
        tree.tag_configure("today", background="#ffd0d0")
        tree.tag_configure("overdue", background="#ffbcbc")
        tree.tag_configure("completed", foreground="#69806D", background="#E8F2E5")
        tree.tag_configure("changed", background="#E7DCF7")
        tree.tag_configure("error", background="#ffd6d6")
        self._enable_column_drag(tree)
        def clear_selection_on_blank(event: tk.Event) -> None:
            if not tree.identify_row(event.y) and tree.identify_region(event.x, event.y) != "heading":
                tree.selection_remove(tree.selection())

        tree.bind("<Button-1>", clear_selection_on_blank, add="+")
        tree.bind("<Configure>", lambda _e: self.after_idle(lambda: self._redraw_progress_if_needed(tree)), add="+")
        tree.bind("<MouseWheel>", lambda _e: self.after(30, lambda: self._redraw_progress_if_needed(tree)), add="+")
        tree.bind("<<TreeviewSelect>>", lambda _e: self.after_idle(lambda: self._redraw_progress_if_needed(tree)), add="+")
        return tree

    def _sort_tree(self, tree: ttk.Treeview, column: str) -> None:
        previous_column = getattr(tree, "_sort_column", None)
        previous_reverse = bool(getattr(tree, "_sort_reverse", False))
        reverse = not previous_reverse if previous_column == column else False

        def sort_key(iid: str) -> tuple[int, object]:
            value = str(tree.set(iid, column)).strip()
            if not value:
                return (2, "")
            if value == "오늘":
                return (0, 0)
            if value.startswith("D-") and value[2:].isdigit():
                return (0, int(value[2:]))
            if value.startswith("D+") and value[2:].isdigit():
                return (0, -int(value[2:]))
            try:
                return (0, float(value.replace(",", "")))
            except ValueError:
                return (1, value.casefold())

        rows = list(tree.get_children(""))
        rows.sort(key=sort_key, reverse=reverse)
        for index, iid in enumerate(rows):
            tree.move(iid, "", index)
        tree._sort_column = column  # type: ignore[attr-defined]
        tree._sort_reverse = reverse  # type: ignore[attr-defined]
        self.after_idle(lambda: self._redraw_progress_if_needed(tree))

    def _redraw_progress_if_needed(self, tree: ttk.Treeview) -> None:
        if getattr(self, "complaint_tree", None) is tree:
            self._draw_complaint_progress()
        elif getattr(self, "task_tree", None) is tree:
            self._draw_task_progress()
        else:
            self._draw_tree_grid(tree)

    @staticmethod
    def _select_tree_row(tree: ttk.Treeview, row_id: str, event: tk.Event) -> None:
        if event.state & 0x0004:
            if row_id in tree.selection():
                tree.selection_remove(row_id)
            else:
                tree.selection_add(row_id)
        elif event.state & 0x0001 and tree.focus():
            rows = list(tree.get_children(""))
            try:
                start = rows.index(tree.focus())
                end = rows.index(row_id)
                tree.selection_set(rows[min(start, end): max(start, end) + 1])
            except ValueError:
                tree.selection_set(row_id)
        else:
            tree.selection_set(row_id)
        tree.focus(row_id)

    def _enable_column_drag(self, tree: ttk.Treeview) -> None:
        drag_state: dict[str, int | None] = {"source": None}

        def display_order() -> list[str]:
            configured = tree.cget("displaycolumns")
            if configured in ("#all", ("#all",)):
                return list(tree.cget("columns"))
            if isinstance(configured, str):
                return list(tree.tk.splitlist(configured))
            return list(configured)

        def press(event: tk.Event) -> None:
            if tree.identify_region(event.x, event.y) != "heading":
                drag_state["source"] = None
                return
            column = tree.identify_column(event.x)
            if column.startswith("#") and column[1:].isdigit():
                drag_state["source"] = int(column[1:]) - 1
                tree.configure(cursor="fleur")

        def release(event: tk.Event) -> None:
            source = drag_state.get("source")
            tree.configure(cursor="")
            drag_state["source"] = None
            if source is None or tree.identify_region(event.x, event.y) != "heading":
                return
            column = tree.identify_column(event.x)
            if not column.startswith("#") or not column[1:].isdigit():
                return
            target = int(column[1:]) - 1
            order = display_order()
            if not (0 <= source < len(order) and 0 <= target < len(order)) or source == target:
                return
            moved = order.pop(source)
            order.insert(target, moved)
            tree.configure(displaycolumns=order)
            self.after_idle(lambda: self._redraw_progress_if_needed(tree))

        tree.bind("<ButtonPress-1>", press, add="+")
        tree.bind("<ButtonRelease-1>", release, add="+")

    def _toolbar(self, parent: tk.Misc) -> ttk.Frame:
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(0, 10))
        return bar

    def _build_today_tab(self) -> None:
        bar = self._toolbar(self.today_tab)
        guidance = ttk.Frame(bar)
        guidance.pack(side="left")
        ttk.Label(guidance, text="※ 기한 초과 및 7일 이내 업무 · 오늘의 일정", style="Subtitle.TLabel").pack(anchor="w")
        ttk.Label(guidance, text="※ 열 제목을 드래그하면 순서를 바꿀 수 있습니다", style="Subtitle.TLabel").pack(anchor="w")
        ttk.Button(bar, text="새로고침", command=self.refresh_all).pack(side="right")
        self.today_tree = self._new_tree(
            self.today_tab,
            ("kind", "title", "person", "deadline", "dday", "status"),
            ("구분", "업무/민원/일정명", "대상자", "날짜/시간", "남은 기간", "상태"),
            (90, 300, 150, 160, 100, 90),
        )
        self.today_tree.tag_configure("personal_event", foreground="#202020", background="#F5FCEF")
        self.today_tree.tag_configure("department_event", foreground="#155AA8", background="#DDEEFF")
        self.today_tree.bind("<Double-1>", self.open_today_item)

    def _build_complaints_tab(self) -> None:
        bar = self._toolbar(self.complaints_tab)
        self.complaint_search = tk.StringVar()
        ttk.Label(bar, text="검색").pack(side="left")
        entry = ttk.Entry(bar, textvariable=self.complaint_search, width=26)
        entry.pack(side="left", padx=(7, 10))
        entry.bind("<KeyRelease>", lambda _e: self.refresh_complaints())
        ttk.Label(bar, text="※ Ctrl+클릭 또는 Shift+클릭으로 여러 항목을 선택할 수 있습니다.", style="Subtitle.TLabel").pack(side="left")
        self.complaint_quick_filter = tk.StringVar(value="전체")
        complaint_filter = ttk.Combobox(bar, textvariable=self.complaint_quick_filter, values=("전체", "오늘 마감", "3일 이내", "이번 주", "기한 초과"), state="readonly", width=12)
        complaint_filter.pack(side="right")
        ttk.Label(bar, text="빠른 필터").pack(side="right", padx=(0, 7))
        complaint_filter.bind("<<ComboboxSelected>>", lambda _event: self.refresh_complaints())
        actions = ttk.Frame(self.complaints_tab)
        actions.pack(fill="x", pady=(0, 10))
        ttk.Button(actions, text="민원 수기 등록", style="Accent.TButton", command=self.add_complaint_dialog).pack(side="left")
        ttk.Button(actions, text="엑셀로 가져오기", command=self.open_excel_import).pack(side="left", padx=6)
        ttk.Button(actions, text="선택 항목 일괄 변경", style="Accent.TButton", command=lambda: self.open_bulk_change_dialog("complaint")).pack(side="left")
        ttk.Button(actions, text="변경 이력", command=lambda: self.show_audit_history("complaint")).pack(side="left", padx=6)
        right_actions = ttk.Frame(actions)
        right_actions.pack(side="right")
        ttk.Button(right_actions, text="진행 현황 사용자 정의", style="Accent.TButton", command=self.open_stage_editor).pack(side="left", padx=6)
        ttk.Button(right_actions, text="선택 항목 삭제", command=lambda: self.move_selected_to_trash("complaint")).pack(side="left", padx=6)
        ttk.Button(right_actions, text="처리 완료", command=self.complete_complaint).pack(side="left", padx=6)
        ttk.Button(right_actions, text="처리기한 연장", command=self.extend_complaint_dialog).pack(side="left", padx=6)
        self.complaint_tree = self._new_tree(
            self.complaints_tab,
            ("receipt", "name", "person", "received", "deadline", "dday", "progress", "progress_edit"),
            ("접수번호", "민원명", "민원인", "받은일자", "현재 처리기한", "남은 기간", "진행 현황", "진행 현황 수정"),
            (110, 170, 90, 95, 130, 75, 410, 120),
        )
        self.complaint_tree.bind("<Double-1>", self.edit_complaint_dialog)

    def _build_tasks_tab(self) -> None:
        bar = self._toolbar(self.tasks_tab)
        self.task_search = tk.StringVar()
        ttk.Label(bar, text="검색").pack(side="left")
        entry = ttk.Entry(bar, textvariable=self.task_search, width=26)
        entry.pack(side="left", padx=(7, 10))
        entry.bind("<KeyRelease>", lambda _e: self.refresh_tasks())
        ttk.Label(bar, text="※ Ctrl+클릭 또는 Shift+클릭으로 여러 항목을 선택할 수 있습니다.", style="Subtitle.TLabel").pack(side="left")
        self.task_quick_filter = tk.StringVar(value="전체")
        task_filter = ttk.Combobox(bar, textvariable=self.task_quick_filter, values=("전체", "오늘 마감", "3일 이내", "이번 주", "기한 초과"), state="readonly", width=12)
        task_filter.pack(side="right")
        ttk.Label(bar, text="빠른 필터").pack(side="right", padx=(0, 7))
        task_filter.bind("<<ComboboxSelected>>", lambda _event: self.refresh_tasks())
        actions = ttk.Frame(self.tasks_tab)
        actions.pack(fill="x", pady=(0, 10))
        ttk.Button(actions, text="수시업무 등록", style="Accent.TButton", command=self.add_task_dialog).pack(side="left")
        ttk.Button(actions, text="엑셀로 가져오기", command=self.open_task_excel_import).pack(side="left", padx=6)
        ttk.Button(actions, text="선택 항목 일괄 변경", style="Accent.TButton", command=lambda: self.open_bulk_change_dialog("task")).pack(side="left", padx=6)
        ttk.Button(actions, text="변경 이력", command=lambda: self.show_audit_history("task")).pack(side="left")
        right_actions = ttk.Frame(actions)
        right_actions.pack(side="right")
        ttk.Button(right_actions, text="진행 현황 사용자 정의", style="Accent.TButton", command=self.open_task_stage_editor).pack(side="left", padx=6)
        ttk.Button(right_actions, text="선택 항목 삭제", command=lambda: self.move_selected_to_trash("task")).pack(side="left", padx=6)
        ttk.Button(right_actions, text="처리 완료", command=self.complete_task).pack(side="left")
        self.task_tree = self._new_tree(
            self.tasks_tab,
            ("title", "registered", "deadline", "dday", "priority", "progress", "progress_edit"),
            ("업무 제목", "등록일", "처리기한", "남은 기간", "비고", "진행 현황", "진행 현황 수정"),
            (280, 110, 140, 90, 90, 390, 120),
        )
        self.task_tree.bind("<Double-1>", self.edit_task_dialog)

    def _build_calendar_tab(self) -> None:
        toolbar = self._toolbar(self.calendar_tab)
        ttk.Button(toolbar, text="◀ 이전 달", command=lambda: self.change_calendar_month(-1)).pack(side="left")
        self.calendar_month_var = tk.StringVar()
        ttk.Label(toolbar, textvariable=self.calendar_month_var, style="Title.TLabel").pack(side="left", padx=18)
        ttk.Button(toolbar, text="다음 달 ▶", command=lambda: self.change_calendar_month(1)).pack(side="left")
        ttk.Button(toolbar, text="오늘", style="Accent.TButton", command=self.go_calendar_today).pack(side="left", padx=8)

        self.calendar_filter = tk.StringVar(value="전체")
        filter_box = ttk.Combobox(
            toolbar,
            textvariable=self.calendar_filter,
            values=("전체", "새올 민원", "수시업무", "개인일정", "부서일정"),
            state="readonly",
            width=12,
        )
        filter_box.pack(side="right")
        ttk.Label(toolbar, text="업무 구분").pack(side="right", padx=(0, 7))
        filter_box.bind("<<ComboboxSelected>>", lambda _event: self.refresh_calendar())

        ttk.Label(
            self.calendar_tab,
            text="날짜를 한 번 누르면 상세 목록을 보고, 날짜를 더블클릭하면 새 일정을 등록할 수 있습니다.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(0, 8))

        content = ttk.Frame(self.calendar_tab)
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.columnconfigure(1, weight=0, minsize=470)
        content.rowconfigure(0, weight=1)
        calendar_panel = ttk.Frame(content)
        calendar_panel.grid(row=0, column=0, sticky="nsew")
        detail_panel = ttk.LabelFrame(content, text="전체 일정 요약", padding=10)
        detail_panel.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        detail_panel.columnconfigure(0, weight=1)
        detail_panel.rowconfigure(1, weight=1)

        weekday_row = ttk.Frame(calendar_panel)
        weekday_row.pack(fill="x")
        for column, weekday in enumerate(("월", "화", "수", "목", "금", "토", "일")):
            foreground = "#3478C5" if column == 5 else "#D74343" if column == 6 else PALETTE["green_dark"]
            label = tk.Label(
                weekday_row,
                text=weekday,
                font=(self.ui_font, 10, "bold"),
                foreground=foreground,
                background=PALETTE["green_soft"],
                pady=6,
            )
            label.grid(row=0, column=column, sticky="ew", padx=1)
            weekday_row.columnconfigure(column, weight=1)

        self.calendar_days_frame = ttk.Frame(calendar_panel)
        self.calendar_days_frame.pack(fill="both", expand=True, pady=(4, 0))
        for column in range(7):
            self.calendar_days_frame.columnconfigure(column, weight=1, uniform="calendar-column")

        self.calendar_selected_var = tk.StringVar()
        summary_header = ttk.Frame(detail_panel)
        summary_header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(summary_header, textvariable=self.calendar_selected_var, style="Subtitle.TLabel", wraplength=260).pack(side="left")
        ttk.Button(summary_header, text="삭제", command=self.delete_selected_calendar_item).pack(side="right")
        ttk.Button(summary_header, text="수정", style="Accent.TButton", command=self.edit_selected_calendar_item).pack(side="right", padx=(0, 6))

        calendar_list_panel = ttk.Frame(detail_panel)
        calendar_list_panel.grid(row=1, column=0, sticky="nsew")
        self.calendar_day_tree = self._new_calendar_tree(
            calendar_list_panel,
            ("kind", "title", "time"),
            ("구분", "제목", "시간"),
            (75, 230, 110),
            height=7,
        )
        self.calendar_day_tree.bind("<<TreeviewSelect>>", lambda _event: self.show_calendar_detail())
        self.calendar_day_tree.bind(
            "<Double-1>",
            lambda event: self.open_calendar_item(self.calendar_day_tree, event),
        )
        self.calendar_day_tree.tag_configure("complaint", foreground="#8A4DBF")
        self.calendar_day_tree.tag_configure("task", foreground="#3478C5")
        self.calendar_day_tree.tag_configure("personal_event", foreground="#202020")
        self.calendar_day_tree.tag_configure("department_event", foreground="#155AA8", background="#DDEEFF")

        self.calendar_detail_frame = ttk.LabelFrame(detail_panel, text="일정 상세", padding=10)
        self.calendar_detail_frame.configure(height=170)
        self.calendar_detail_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        self.calendar_detail_frame.pack_propagate(False)
        self.calendar_detail_title_var = tk.StringVar(value="목록에서 항목을 선택하세요.")
        self.calendar_detail_var = tk.StringVar(value="")
        ttk.Label(self.calendar_detail_frame, textvariable=self.calendar_detail_title_var).pack(anchor="w")
        ttk.Label(
            self.calendar_detail_frame,
            textvariable=self.calendar_detail_var,
            style="Subtitle.TLabel",
            wraplength=420,
            justify="left",
        ).pack(anchor="w", fill="x", pady=(6, 0))

    def _new_calendar_tree(
        self,
        parent: tk.Misc,
        columns: tuple[str, ...],
        headings: tuple[str, ...],
        widths: tuple[int, ...],
        height: int,
    ) -> ttk.Treeview:
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)
        tree = ttk.Treeview(
            container,
            columns=columns,
            show="headings",
            style="Calendar.Treeview",
            selectmode="browse",
            height=height,
        )
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        for column, heading, width in zip(columns, headings, widths):
            tree.heading(column, text=heading)
            tree.column(column, width=width, minwidth=60, anchor="center" if column in {"kind", "deadline", "dday"} else "w")
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return tree

    def _build_completed_tab(self) -> None:
        bar = self._toolbar(self.completed_tab)
        guidance = ttk.Frame(bar)
        guidance.pack(side="left")
        ttk.Label(guidance, text="※ 완료 내역은 최근 30개까지만 저장됩니다.", style="Subtitle.TLabel").pack(anchor="w")
        ttk.Button(bar, text="선택 항목 복원", style="Accent.TButton", command=self.restore_completed_item).pack(side="right")
        self.completed_tree = self._new_tree(
            self.completed_tab,
            ("kind", "title", "person", "deadline", "completed"),
            ("구분", "업무/민원명", "대상자", "처리기한", "완료일"),
            (90, 340, 160, 170, 170),
        )
        self.completed_tree.bind("<Double-1>", lambda _e: self.restore_completed_item())

    def _build_trash_tab(self) -> None:
        bar = self._toolbar(self.trash_tab)
        guidance = ttk.Frame(bar)
        guidance.pack(side="left")
        ttk.Label(guidance, text="※ 삭제한 항목은 복원할 수 있습니다.", style="Subtitle.TLabel").pack(anchor="w")
        ttk.Label(guidance, text="※ 영구 삭제한 항목은 복구할 수 없습니다.", style="Subtitle.TLabel").pack(anchor="w")
        ttk.Button(bar, text="선택 항목 영구 삭제", command=self.permanently_delete_trash).pack(side="right")
        ttk.Button(bar, text="선택 항목 복원", style="Accent.TButton", command=self.restore_trash_items).pack(side="right", padx=6)
        self.trash_tree = self._new_tree(
            self.trash_tab,
            ("kind", "title", "person", "deadline", "deleted"),
            ("구분", "업무/민원명", "대상자", "처리기한", "삭제일"),
            (90, 360, 160, 170, 180),
        )
        self.trash_tree.bind("<Double-1>", lambda _event: self.restore_trash_items())

    def add_complaint_dialog(self) -> None:
        today = date.today().isoformat()
        fields = [
            ("receipt_no", "접수번호 *", ""), ("complaint_name", "민원명 *", ""),
            ("applicant_name", "민원인 성명", ""),
            ("birth_date", "생년월일", ""), ("received_at", "받은일자 *", today),
            ("registered_at", "접수일자", today), ("deadline", f"처리기한 * ({DATE_HINT})", today),
            ("memo", "메모", ""),
        ]

        def save(data: dict[str, str]) -> None:
            for key, label in (("receipt_no", "접수번호"), ("complaint_name", "민원명"), ("received_at", "받은일자"), ("deadline", "처리기한")):
                if not data[key]:
                    raise ValueError(f"{label}을 입력하세요.")
            parse_deadline(data["received_at"])
            parse_deadline(data["deadline"])
            data["processing_stage"] = self._stage_label(0)
            self.db.add_complaint(data)
            self.refresh_all()

        FormDialog(self, "민원 수기 등록", fields, save)

    def add_task_dialog(self) -> None:
        today = date.today().isoformat()
        fields = [
            ("title", "업무 제목 *", ""), ("registered_at", "등록일 *", today),
            ("deadline", f"처리기한 * ({DATE_HINT})", today), ("priority", "비고", ""),
            ("content", "업무 내용", ""),
        ]

        def save(data: dict[str, str]) -> None:
            if not data["title"] or not data["registered_at"] or not data["deadline"]:
                raise ValueError("업무 제목, 등록일, 처리기한을 입력하세요.")
            parse_deadline(data["registered_at"])
            parse_deadline(data["deadline"])
            data["processing_stage"] = str(self.task_stage_config[0]["name"])
            self.db.add_task(data)
            self.refresh_all()

        FormDialog(self, "수시업무 등록", fields, save)

    def _selected_id(self, tree: ttk.Treeview) -> int | None:
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("항목 선택", "먼저 목록에서 항목을 선택하세요.", parent=self)
            return None
        return int(selected[0])

    def edit_complaint_dialog(self, event: tk.Event | None = None) -> None:
        if event is not None:
            row_id = self.complaint_tree.identify_row(event.y)
            if not row_id:
                return
            self.complaint_tree.selection_set(row_id)
        complaint_id = self._selected_id(self.complaint_tree)
        if complaint_id is None:
            return
        complaint = self.db.complaint_by_id(complaint_id)
        if not complaint:
            return

        window = tk.Toplevel(self)
        window.title("민원 신청 내용 수정")
        window.transient(self)
        window.geometry("720x590")
        window.resizable(True, True)
        window.minsize(600, 480)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="민원 신청 내용 수정", style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))
        body.columnconfigure(1, weight=1)

        field_specs = (
            ("receipt_no", "접수번호 *"), ("complaint_name", "민원명 *"),
            ("applicant_name", "민원인"), ("birth_date", "생년월일"),
            ("registered_at", "접수일자 *"), ("current_deadline", "현재 처리기한 *"),
        )
        variables: dict[str, tk.StringVar] = {}
        row = 1
        for key, label in field_specs:
            ttk.Label(body, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            variable = tk.StringVar(value=complaint.get(key, ""))
            ttk.Entry(body, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=5)
            variables[key] = variable
            row += 1

        ttk.Label(body, text="메모").grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=5)
        memo = tk.Text(body, height=5, wrap="word", font=self.input_font)
        memo.insert("1.0", complaint.get("memo", ""))
        memo.grid(row=row, column=1, sticky="nsew", pady=5)
        body.rowconfigure(row, weight=1)
        row += 1

        buttons = ttk.Frame(body)
        buttons.grid(row=row, column=0, columnspan=2, sticky="e", pady=(14, 0))

        def save() -> None:
            data = {key: variable.get().strip() for key, variable in variables.items()}
            for key, label in (("receipt_no", "접수번호"), ("complaint_name", "민원명"), ("registered_at", "접수일자"), ("current_deadline", "현재 처리기한")):
                if not data[key]:
                    messagebox.showwarning("입력 확인", f"{label}을 입력하세요.", parent=window)
                    return
            try:
                parse_deadline(data["registered_at"])
                parse_deadline(data["current_deadline"])
            except ValueError as exc:
                messagebox.showwarning("날짜 확인", str(exc), parent=window)
                return
            data["deadline"] = data.pop("current_deadline")
            data["received_at"] = complaint["received_at"]
            data["intake_type"] = complaint["intake_type"]
            data["assignee"] = complaint["assignee"]
            data["processing_stage"] = complaint["processing_stage"]
            data["memo"] = memo.get("1.0", "end").strip()
            try:
                self.db.update_complaint(complaint_id, data)
            except Exception as exc:
                messagebox.showerror("수정할 수 없음", str(exc), parent=window)
                return
            self.refresh_all()
            window.destroy()

        ttk.Button(buttons, text="취소", command=window.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="수정 저장", style="Accent.TButton", command=save).pack(side="right")

    def edit_task_dialog(self, event: tk.Event | None = None) -> None:
        if event is not None:
            row_id = self.task_tree.identify_row(event.y)
            if not row_id:
                return
            self.task_tree.selection_set(row_id)
        task_id = self._selected_id(self.task_tree)
        if task_id is None:
            return
        task = self.db.task_by_id(task_id)
        if not task:
            return
        fields = [
            ("title", "업무 제목 *", task["title"]),
            ("registered_at", "등록일 *", task["registered_at"]),
            ("deadline", f"처리기한 * ({DATE_HINT})", task["deadline"]),
            ("priority", "비고", task["priority"]),
            ("content", "업무 내용", task["content"]),
        ]

        def save(data: dict[str, str]) -> None:
            if not data["title"] or not data["registered_at"] or not data["deadline"]:
                raise ValueError("업무 제목, 등록일, 처리기한을 입력하세요.")
            parse_deadline(data["registered_at"])
            parse_deadline(data["deadline"])
            self.db.update_task(task_id, data)
            self.refresh_all()

        FormDialog(self, "수시업무 내용 수정", fields, save)

    def open_item_progress_editor(self, kind: str, item_id: int) -> None:
        if kind == "complaint":
            item = self.db.complaint_by_id(item_id)
            if not item:
                return
            nodes = self._progress_node_keys()
            state = self._progress_state(item)
            node_text = self._definition_text_from_key
        else:
            item = self.db.task_by_id(item_id)
            if not item:
                return
            stages = self._task_stages(item)
            nodes = self._progress_node_keys(stages)
            state = self._task_progress_state(item)
            node_text = lambda key: self._task_stage_text(key, stages)

        completed = set(state.get("completed", []))
        window = tk.Toplevel(self)
        window.title("진행 현황 수정")
        window.transient(self)
        window.geometry("620x600")
        window.resizable(True, True)
        window.minsize(500, 420)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="진행 현황 수정", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text="완료된 단계에 체크하세요. 뒤 단계를 체크하면 앞 단계도 함께 완료됩니다.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        list_frame = ttk.Frame(body)
        list_frame.pack(fill="both", expand=True)
        variables: list[tk.BooleanVar] = []

        def sync_checks(index: int) -> None:
            if variables[index].get():
                for previous in variables[: index + 1]:
                    previous.set(True)
            else:
                for following in variables[index:]:
                    following.set(False)

        for index, node in enumerate(nodes):
            variable = tk.BooleanVar(value=node in completed)
            variables.append(variable)
            row = ttk.Frame(list_frame, padding=(10, 8))
            row.pack(fill="x", pady=3)
            ttk.Checkbutton(
                row,
                text=f"{index + 1}. {node_text(node)}",
                variable=variable,
                command=lambda selected=index: sync_checks(selected),
            ).pack(anchor="w")

        def save_progress() -> None:
            checked_nodes = [node for node, variable in zip(nodes, variables) if variable.get()]
            pending = [node for node in nodes if node not in checked_nodes]
            current = pending[0] if pending else nodes[-1]
            new_state = {"current": current, "completed": checked_nodes}
            if kind == "complaint":
                self.db.update_progress_state(item_id, new_state, node_text(current))
            else:
                self.db.update_task_progress_state(item_id, new_state, node_text(current))
            self.refresh_all()
            window.destroy()

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(14, 0))
        ttk.Button(actions, text="취소", command=window.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(actions, text="수정 저장", style="Accent.TButton", command=save_progress).pack(side="right")

    def open_progress_editor(self) -> None:
        complaint_id = self._selected_id(self.complaint_tree)
        if complaint_id is None:
            return
        complaint = self.db.complaint_by_id(complaint_id)
        if not complaint:
            return
        state = self._progress_state(complaint)
        completed_keys = set(state.get("completed", []))

        window = tk.Toplevel(self)
        window.title("진행 현황 설정")
        window.transient(self)
        window.geometry("760x680")
        window.resizable(True, True)
        window.minsize(620, 500)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="진행 현황 설정", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text=f"{complaint['complaint_name']} · 완료 여부와 현재 단계를 선택한 뒤 저장하세요.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 10))

        legend = ttk.Frame(body)
        legend.pack(fill="x", pady=(0, 10))
        for label, color in (("완료", PROGRESS_COLORS["completed"]), ("현재 단계", PROGRESS_COLORS["current"]), ("실행 전", PROGRESS_COLORS["pending"])):
            swatch = tk.Label(legend, text="  ", background=color, relief="flat")
            swatch.pack(side="left", padx=(0, 4))
            ttk.Label(legend, text=label).pack(side="left", padx=(0, 16))

        container = ttk.Frame(body)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, background=PALETTE["background"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        form = ttk.Frame(canvas)
        form.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=form, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        current_var = tk.StringVar(value=str(state.get("current", "")))
        completed_vars: dict[str, tk.BooleanVar] = {}

        for stage_index, stage in enumerate(self.stage_config):
            stage_key = f"s:{stage['id']}"
            stage_panel = ttk.LabelFrame(form, text=f"{stage_index + 1}. {stage['name']}", padding=10)
            stage_panel.pack(fill="x", padx=4, pady=5)
            completed_var = tk.BooleanVar(value=stage_key in completed_keys)
            completed_vars[stage_key] = completed_var
            ttk.Checkbutton(stage_panel, text="완료", variable=completed_var).pack(side="left")
            ttk.Radiobutton(stage_panel, text="현재 단계", variable=current_var, value=stage_key).pack(side="left", padx=(18, 0))
            branches = stage.get("branches", [])
            if branches:
                branch_frame = ttk.Frame(form)
                branch_frame.pack(fill="x", padx=(40, 8), pady=(0, 5))
                for branch_index, branch in enumerate(branches):
                    branch_key = f"b:{stage['id']}:{branch['id']}"
                    branch_row = ttk.Frame(branch_frame)
                    branch_row.pack(fill="x", pady=2)
                    ttk.Label(branch_row, text=f"↳ {stage_index + 1}-{branch_index + 1}. {branch['name']}", width=34).pack(side="left")
                    branch_completed = tk.BooleanVar(value=branch_key in completed_keys)
                    completed_vars[branch_key] = branch_completed
                    ttk.Checkbutton(branch_row, text="완료", variable=branch_completed).pack(side="left")
                    ttk.Radiobutton(branch_row, text="현재 분기", variable=current_var, value=branch_key).pack(side="left", padx=(18, 0))

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(14, 0))

        def stage_text_from_key(key: str) -> str:
            parts = key.split(":")
            if len(parts) < 2:
                return self._stage_label(0)
            stage_index = next((index for index, stage in enumerate(self.stage_config) if str(stage["id"]) == parts[1]), 0)
            text = self._stage_label(stage_index)
            if parts[0] == "b" and len(parts) == 3:
                branches = self.stage_config[stage_index].get("branches", [])
                branch_index = next((index for index, branch in enumerate(branches) if str(branch["id"]) == parts[2]), -1)
                if branch_index >= 0:
                    text += f" > {stage_index + 1}-{branch_index + 1}. {branches[branch_index]['name']}"
            return text

        def save_progress() -> None:
            current_key = current_var.get()
            if not current_key:
                messagebox.showwarning("현재 단계", "현재 단계 또는 현재 분기를 하나 선택하세요.", parent=window)
                return
            completed = [key for key, variable in completed_vars.items() if variable.get() and key != current_key]
            new_state = {"current": current_key, "completed": completed}
            self.db.update_progress_state(complaint_id, new_state, stage_text_from_key(current_key))
            self.refresh_all()
            window.destroy()

        ttk.Button(buttons, text="취소", command=window.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="저장", style="Accent.TButton", command=save_progress).pack(side="right")

    def open_stage_update_dialog(self, event: tk.Event | None = None) -> None:
        if event is not None:
            row_id = self.complaint_tree.identify_row(event.y)
            if not row_id:
                return
            self.complaint_tree.selection_set(row_id)
        complaint_id = self._selected_id(self.complaint_tree)
        if complaint_id is None:
            return
        complaint = self.db.complaint_by_id(complaint_id)
        if not complaint:
            return

        window = tk.Toplevel(self)
        window.title("현재 처리단계 기록")
        window.transient(self)
        window.geometry("620x540")
        window.resizable(True, True)
        window.minsize(520, 420)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=complaint["complaint_name"], style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text=f"접수번호 {complaint['receipt_no']} · {complaint['applicant_name']}",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 16))

        ttk.Label(body, text="주요 단계").pack(anchor="w")
        stage_var = tk.StringVar()
        stage_box = ttk.Combobox(
            body,
            textvariable=stage_var,
            values=[self._stage_label(index) for index in range(len(self.stage_config))],
            state="readonly",
        )
        stage_box.pack(fill="x", pady=(4, 12))

        ttk.Label(body, text="세부 분기").pack(anchor="w")
        branch_var = tk.StringVar()
        branch_box = ttk.Combobox(body, textvariable=branch_var, state="readonly")
        branch_box.pack(fill="x", pady=(4, 14))

        current_index = self._stage_index(complaint["processing_stage"])
        stage_box.current(current_index if current_index >= 0 else 0)

        def refresh_branches(_event: object | None = None) -> None:
            selected_index = stage_box.current()
            branches = self.stage_config[selected_index].get("branches", []) if selected_index >= 0 else []
            labels = [f"{selected_index + 1}-{index + 1}. {branch['name']}" for index, branch in enumerate(branches)]
            branch_box.configure(values=["선택하지 않음", *labels])
            matched = next(
                (index + 1 for index, branch in enumerate(branches) if str(branch["name"]) in complaint["processing_stage"]),
                0,
            )
            branch_box.current(matched)

        stage_box.bind("<<ComboboxSelected>>", refresh_branches)
        refresh_branches()

        ttk.Label(body, text="단계 변경 이력", style="Subtitle.TLabel").pack(anchor="w", pady=(4, 5))
        history_tree = ttk.Treeview(body, columns=("stage", "time"), show="headings", height=7)
        history_tree.heading("stage", text="변경된 단계")
        history_tree.heading("time", text="변경 일시")
        history_tree.column("stage", width=360)
        history_tree.column("time", width=170, anchor="center")
        history_tree.pack(fill="both", expand=True)
        for item in self.db.processing_stage_history(complaint_id):
            history_tree.insert("", "end", values=(item["new_stage"], item["changed_at"]))

        buttons = ttk.Frame(body)
        buttons.pack(fill="x", pady=(14, 0))

        def save_stage() -> None:
            index = stage_box.current()
            if index < 0:
                messagebox.showwarning("처리단계", "주요 단계를 선택하세요.", parent=window)
                return
            new_stage = self._stage_label(index)
            if branch_box.current() > 0:
                new_stage += " > " + branch_var.get()
            self.db.update_processing_stage(complaint_id, new_stage)
            self.refresh_all()
            window.destroy()

        ttk.Button(buttons, text="취소", command=window.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(buttons, text="현재 단계 저장", style="Accent.TButton", command=save_stage).pack(side="right")

    def _open_stage_editor_legacy(self) -> None:
        working = json.loads(json.dumps(self.stage_config, ensure_ascii=False))
        window = tk.Toplevel(self)
        window.title("진행 현황 사용자 정의")
        window.transient(self)
        window.geometry("820x650")
        window.resizable(True, True)
        window.minsize(700, 550)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="진행 현황 사용자 정의", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text="목록에서 단계를 선택한 다음, 아래 이름 칸을 고치고 '이름 수정 적용'을 누르세요.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        edit_box = ttk.LabelFrame(body, text="선택한 단계·세부 분기 이름 수정", padding=10)
        edit_box.pack(fill="x", pady=(0, 12))
        ttk.Label(edit_box, text="① 목록에서 항목 선택  ② 이름 입력  ③ 이름 수정 적용").pack(anchor="w", pady=(0, 7))
        edit_row = ttk.Frame(edit_box)
        edit_row.pack(fill="x")
        ttk.Label(edit_row, text="바꿀 이름").pack(side="left", padx=(0, 8))
        edit_var = tk.StringVar()
        edit_entry = ttk.Entry(edit_row, textvariable=edit_var, font=self.input_font)
        edit_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(edit_row, text="이름 수정 적용", style="Accent.TButton", command=lambda: apply_name_edit()).pack(side="left", padx=(8, 0))

        tree = ttk.Treeview(body, columns=("name",), show="tree headings", selectmode="browse")
        tree.heading("#0", text="구분")
        tree.heading("name", text="단계 이름")
        tree.column("#0", width=130, anchor="center")
        tree.column("name", width=470)
        tree.pack(fill="both", expand=True)

        def redraw() -> None:
            self._clear_tree(tree)
            for stage_index, stage in enumerate(working):
                stage_iid = f"s:{stage_index}"
                tree.insert("", "end", iid=stage_iid, text=f"{stage_index + 1}단계", values=(stage["name"],), open=True)
                for branch_index, branch in enumerate(stage.get("branches", [])):
                    tree.insert(
                        stage_iid,
                        "end",
                        iid=f"b:{stage_index}:{branch_index}",
                        text=f"{stage_index + 1}-{branch_index + 1}",
                        values=(branch["name"],),
                    )

        def selected_parts() -> list[str] | None:
            selected = tree.selection()
            return selected[0].split(":") if selected else None

        def load_selected_name(_event: object | None = None) -> None:
            parts = selected_parts()
            if not parts:
                edit_var.set("")
                return
            stage_index = int(parts[1])
            value = working[stage_index]["name"] if parts[0] == "s" else working[stage_index]["branches"][int(parts[2])]["name"]
            edit_var.set(str(value))

        def apply_name_edit(_event: object | None = None) -> None:
            parts = selected_parts()
            name = edit_var.get().strip()
            if not parts or not name:
                messagebox.showinfo("이름 수정", "수정할 항목을 선택하고 이름을 입력하세요.", parent=window)
                return
            selected_iid = tree.selection()[0]
            stage_index = int(parts[1])
            if parts[0] == "s":
                working[stage_index]["name"] = name
            else:
                working[stage_index]["branches"][int(parts[2])]["name"] = name
            redraw()
            tree.selection_set(selected_iid)
            tree.see(selected_iid)
            edit_var.set(name)

        def add_stage() -> None:
            working.append({"id": uuid.uuid4().hex, "name": "새 단계", "branches": []})
            redraw()
            tree.selection_set(f"s:{len(working) - 1}")
            load_selected_name()
            edit_entry.focus_set()
            edit_entry.selection_range(0, "end")

        def add_branch() -> None:
            parts = selected_parts()
            if not parts:
                messagebox.showinfo("단계 선택", "세부 분기를 추가할 주요 단계를 선택하세요.", parent=window)
                return
            stage_index = int(parts[1])
            branches = working[stage_index].setdefault("branches", [])
            branches.append({"id": uuid.uuid4().hex, "name": "새 세부 분기"})
            redraw()
            tree.selection_set(f"b:{stage_index}:{len(branches) - 1}")
            load_selected_name()
            edit_entry.focus_set()
            edit_entry.selection_range(0, "end")

        def rename_item() -> None:
            parts = selected_parts()
            if not parts:
                return
            stage_index = int(parts[1])
            if parts[0] == "s":
                current = str(working[stage_index]["name"])
            else:
                current = str(working[stage_index]["branches"][int(parts[2])]["name"])
            name = simpledialog.askstring("이름 변경", "새 이름", initialvalue=current, parent=window)
            if name and name.strip():
                if parts[0] == "s":
                    working[stage_index]["name"] = name.strip()
                else:
                    working[stage_index]["branches"][int(parts[2])]["name"] = name.strip()
                redraw()

        def delete_item() -> None:
            parts = selected_parts()
            if not parts:
                return
            stage_index = int(parts[1])
            if parts[0] == "s":
                if len(working) == 1:
                    messagebox.showwarning("삭제할 수 없음", "주요 단계는 최소 1개가 필요합니다.", parent=window)
                    return
                working.pop(stage_index)
            else:
                working[stage_index]["branches"].pop(int(parts[2]))
            redraw()

        def move(direction: int) -> None:
            parts = selected_parts()
            if not parts:
                return
            stage_index = int(parts[1])
            if parts[0] == "s":
                target = stage_index + direction
                if 0 <= target < len(working):
                    working[stage_index], working[target] = working[target], working[stage_index]
            else:
                branches = working[stage_index]["branches"]
                branch_index = int(parts[2])
                target = branch_index + direction
                if 0 <= target < len(branches):
                    branches[branch_index], branches[target] = branches[target], branches[branch_index]
            redraw()

        toolbar = ttk.Frame(body)
        toolbar.pack(fill="x", pady=(10, 0))
        for text, command in (
            ("단계 추가", add_stage), ("세부 분기 추가", add_branch),
            ("삭제", delete_item), ("위로", lambda: move(-1)), ("아래로", lambda: move(1)),
        ):
            ttk.Button(toolbar, text=text, command=command).pack(side="left", padx=(0, 5))

        def save_config() -> None:
            self.stage_config = working
            self._save_stage_config()
            self.refresh_all()
            window.destroy()

        ttk.Button(toolbar, text="저장", style="Accent.TButton", command=save_config).pack(side="right")
        ttk.Button(toolbar, text="취소", command=window.destroy).pack(side="right", padx=(0, 6))
        tree.bind("<<TreeviewSelect>>", load_selected_name)
        tree.bind("<Double-1>", lambda _event: edit_entry.focus_set())
        edit_entry.bind("<Return>", apply_name_edit)
        redraw()

    def _open_task_stage_editor_legacy(self) -> None:
        working = json.loads(json.dumps(self.task_stage_config, ensure_ascii=False))
        window = tk.Toplevel(self)
        window.title("수시업무 진행 현황 사용자 정의")
        window.transient(self)
        window.geometry("760x620")
        window.resizable(True, True)
        window.minsize(650, 520)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="수시업무 진행 현황 사용자 정의", style="Title.TLabel").pack(anchor="w")
        ttk.Label(body, text="목록에서 단계를 선택한 다음, 아래 이름 칸을 고치고 '이름 수정 적용'을 누르세요.", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 12))

        edit_box = ttk.LabelFrame(body, text="선택한 단계 이름 수정", padding=10)
        edit_box.pack(fill="x", pady=(0, 12))
        ttk.Label(edit_box, text="① 목록에서 단계 선택  ② 이름 입력  ③ 이름 수정 적용").pack(anchor="w", pady=(0, 7))
        edit_row = ttk.Frame(edit_box)
        edit_row.pack(fill="x")
        ttk.Label(edit_row, text="바꿀 이름").pack(side="left", padx=(0, 8))
        edit_var = tk.StringVar()
        edit_entry = ttk.Entry(edit_row, textvariable=edit_var, font=self.input_font)
        edit_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(edit_row, text="이름 수정 적용", style="Accent.TButton", command=lambda: apply_name_edit()).pack(side="left", padx=(8, 0))

        tree = ttk.Treeview(body, columns=("name",), show="headings", selectmode="browse")
        tree.heading("name", text="단계 이름")
        tree.column("name", width=520)
        tree.pack(fill="both", expand=True)

        def redraw() -> None:
            self._clear_tree(tree)
            for index, stage in enumerate(working):
                tree.insert("", "end", iid=str(index), values=(f"{index + 1}. {stage['name']}",))

        def selected_index() -> int | None:
            selected = tree.selection()
            return int(selected[0]) if selected else None

        def load_selected_name(_event: object | None = None) -> None:
            index = selected_index()
            edit_var.set(str(working[index]["name"]) if index is not None else "")

        def apply_name_edit(_event: object | None = None) -> None:
            index = selected_index()
            name = edit_var.get().strip()
            if index is None or not name:
                messagebox.showinfo("이름 수정", "수정할 단계를 선택하고 이름을 입력하세요.", parent=window)
                return
            working[index]["name"] = name
            redraw()
            tree.selection_set(str(index))
            edit_var.set(name)

        def add_stage() -> None:
            working.append({"id": uuid.uuid4().hex, "name": "새 단계"})
            redraw()
            tree.selection_set(str(len(working) - 1))
            load_selected_name()
            edit_entry.focus_set()
            edit_entry.selection_range(0, "end")

        def rename_stage() -> None:
            index = selected_index()
            if index is None:
                return
            name = simpledialog.askstring("이름 변경", "새 단계 이름", initialvalue=working[index]["name"], parent=window)
            if name and name.strip():
                working[index]["name"] = name.strip()
                redraw()

        def delete_stage() -> None:
            index = selected_index()
            if index is None:
                return
            if len(working) == 1:
                messagebox.showwarning("삭제할 수 없음", "단계는 최소 1개가 필요합니다.", parent=window)
                return
            working.pop(index)
            redraw()

        def move(direction: int) -> None:
            index = selected_index()
            if index is None:
                return
            target = index + direction
            if 0 <= target < len(working):
                working[index], working[target] = working[target], working[index]
                redraw()
                tree.selection_set(str(target))

        toolbar = ttk.Frame(body)
        toolbar.pack(fill="x", pady=(10, 0))
        ttk.Button(toolbar, text="단계 추가", command=add_stage).pack(side="left")
        ttk.Button(toolbar, text="삭제", command=delete_stage).pack(side="left")
        ttk.Button(toolbar, text="위로", command=lambda: move(-1)).pack(side="left", padx=(12, 5))
        ttk.Button(toolbar, text="아래로", command=lambda: move(1)).pack(side="left")

        def save() -> None:
            self.task_stage_config = working
            self._save_task_stage_config()
            self.refresh_all()
            window.destroy()

        ttk.Button(toolbar, text="저장", style="Accent.TButton", command=save).pack(side="right")
        ttk.Button(toolbar, text="취소", command=window.destroy).pack(side="right", padx=(0, 6))
        tree.bind("<<TreeviewSelect>>", load_selected_name)
        tree.bind("<Double-1>", lambda _event: edit_entry.focus_set())
        edit_entry.bind("<Return>", apply_name_edit)
        redraw()

    def open_stage_editor(self) -> None:
        self._open_inline_stage_editor(task_mode=False)

    def open_task_stage_editor(self) -> None:
        self._open_inline_stage_editor(task_mode=True)

    def open_item_stage_editor(self, kind: str, item_id: int) -> None:
        if kind == "complaint":
            item = self.db.complaint_by_id(item_id)
            if not item:
                return
            config = self._complaint_stages(item)
        else:
            item = self.db.task_by_id(item_id)
            if not item:
                return
            config = self._task_stages(item)
        self._open_inline_stage_editor(
            task_mode=kind == "task",
            item_kind=kind,
            item_id=item_id,
            initial_config=config,
        )

    def _open_inline_stage_editor(
        self,
        task_mode: bool,
        item_kind: str | None = None,
        item_id: int | None = None,
        initial_config: list[dict[str, object]] | None = None,
    ) -> None:
        config = initial_config or (self.task_stage_config if task_mode else self.stage_config)
        working = json.loads(json.dumps(config, ensure_ascii=False))
        window = tk.Toplevel(self)
        if item_kind:
            window.title("진행 현황 수정")
        else:
            window.title("수시업무 진행 현황 사용자 정의" if task_mode else "진행 현황 사용자 정의")
        window.transient(self)
        window.geometry("840x700")
        window.resizable(True, True)
        window.minsize(680, 520)
        window.grab_set()

        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=window.title(), style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text=("이 단계 구성은 선택한 업무에만 적용되며 변경 내용은 자동 저장됩니다."
                  if item_kind else "각 단계의 글상자를 바로 편집할 수 있습니다. 변경 내용은 자동 저장됩니다."),
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        container = ttk.Frame(body)
        container.pack(fill="both", expand=True)
        canvas = tk.Canvas(container, background=PALETTE["background"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        stage_list = ttk.Frame(canvas)
        list_window = canvas.create_window((0, 0), window=stage_list, anchor="nw")
        stage_list.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(list_window, width=event.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def persist() -> None:
            if item_kind and item_id is not None:
                self.db.update_item_stage_config(item_kind, item_id, working)
            elif task_mode:
                self.task_stage_config = working
                self._save_task_stage_config()
            else:
                self.stage_config = working
                self._save_stage_config()

        def update_name(item: dict[str, object], variable: tk.StringVar) -> None:
            item["name"] = variable.get()
            persist()

        def add_stage() -> None:
            stage: dict[str, object] = {"id": uuid.uuid4().hex, "name": "새 단계", "branches": []}
            working.append(stage)
            persist()
            redraw()
            canvas.yview_moveto(1.0)

        def delete_stage(index: int) -> None:
            if len(working) <= 1:
                messagebox.showwarning("삭제할 수 없음", "단계는 최소 1개가 필요합니다.", parent=window)
                return
            working.pop(index)
            persist()
            redraw()

        def move_stage(index: int, direction: int) -> None:
            target = index + direction
            if 0 <= target < len(working):
                working[index], working[target] = working[target], working[index]
                persist()
                redraw()

        def add_branch(stage_index: int) -> None:
            branches = working[stage_index].setdefault("branches", [])
            branches.append({"id": uuid.uuid4().hex, "name": "새 세부 분기"})
            persist()
            redraw()

        def delete_branch(stage_index: int, branch_index: int) -> None:
            working[stage_index].setdefault("branches", []).pop(branch_index)
            persist()
            redraw()

        def redraw() -> None:
            for child in stage_list.winfo_children():
                child.destroy()
            for stage_index, stage in enumerate(working):
                panel = ttk.LabelFrame(stage_list, text=f"{stage_index + 1}단계", padding=10)
                panel.pack(fill="x", padx=4, pady=6)
                header = ttk.Frame(panel)
                header.pack(fill="x")
                ttk.Label(header, text="단계 이름").pack(side="left", padx=(0, 8))
                stage_var = tk.StringVar(value=str(stage.get("name", "")))
                stage_entry = ttk.Entry(header, textvariable=stage_var, font=self.input_font)
                stage_entry.pack(side="left", fill="x", expand=True)
                stage_var.trace_add("write", lambda *_args, item=stage, variable=stage_var: update_name(item, variable))
                ttk.Button(header, text="단계 삭제", command=lambda index=stage_index: delete_stage(index)).pack(side="right", padx=(8, 0))
                ttk.Button(header, text="▼", width=3, command=lambda index=stage_index: move_stage(index, 1)).pack(side="right", padx=(4, 0))
                ttk.Button(header, text="▲", width=3, command=lambda index=stage_index: move_stage(index, -1)).pack(side="right", padx=(8, 0))

                branches = stage.setdefault("branches", [])
                for branch_index, branch in enumerate(branches):
                    branch_row = ttk.Frame(panel)
                    branch_row.pack(fill="x", padx=(36, 0), pady=(7, 0))
                    ttk.Label(branch_row, text=f"{stage_index + 1}-{branch_index + 1}").pack(side="left", padx=(0, 8))
                    branch_var = tk.StringVar(value=str(branch.get("name", "")))
                    branch_entry = ttk.Entry(branch_row, textvariable=branch_var, font=self.input_font)
                    branch_entry.pack(side="left", fill="x", expand=True)
                    branch_var.trace_add("write", lambda *_args, item=branch, variable=branch_var: update_name(item, variable))
                    ttk.Button(
                        branch_row,
                        text="분기 삭제",
                        command=lambda si=stage_index, bi=branch_index: delete_branch(si, bi),
                    ).pack(side="right", padx=(8, 0))
                ttk.Button(panel, text="+ 세부 분기 추가", command=lambda index=stage_index: add_branch(index)).pack(anchor="w", padx=(36, 0), pady=(9, 0))

        def close_editor() -> None:
            persist()
            self.refresh_all()
            window.destroy()

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(12, 0))
        ttk.Button(actions, text="+ 단계 추가", style="Accent.TButton", command=add_stage).pack(side="left")
        ttk.Button(actions, text="닫기", style="Accent.TButton", command=close_editor).pack(side="right")
        window.protocol("WM_DELETE_WINDOW", close_editor)
        redraw()

    def extend_complaint_dialog(self) -> None:
        item_id = self._selected_id(self.complaint_tree)
        if item_id is None:
            return
        values = self.complaint_tree.item(str(item_id), "values")
        old_deadline = values[4] if values else ""
        fields = [("deadline", f"새 처리기한 * ({DATE_HINT})", old_deadline), ("reason", "연장 사유 *", "")]

        def save(data: dict[str, str]) -> None:
            parse_deadline(data["deadline"])
            if not data["reason"]:
                raise ValueError("연장 사유를 입력하세요.")
            if data["deadline"] == old_deadline:
                raise ValueError("기존 처리기한과 다른 날짜를 입력하세요.")
            self.db.extend_deadline(item_id, data["deadline"], data["reason"])
            self.refresh_all()

        FormDialog(self, "처리기한 연장", fields, save)

    def open_bulk_change_dialog(self, kind: str) -> None:
        tree = self.complaint_tree if kind == "complaint" else self.task_tree
        item_ids = [int(iid) for iid in tree.selection()]
        if not item_ids:
            messagebox.showinfo("항목 선택", "일괄 변경할 항목을 먼저 선택하세요.", parent=self)
            return
        items = [self.db.complaint_by_id(item_id) if kind == "complaint" else self.db.task_by_id(item_id) for item_id in item_ids]
        items = [item for item in items if item]
        max_nodes = max(
            len(self._progress_node_keys(self._complaint_stages(item))) if kind == "complaint" else len(self._progress_node_keys(self._task_stages(item)))
            for item in items
        )

        window = tk.Toplevel(self)
        window.title("선택 항목 일괄 변경")
        window.transient(self)
        window.geometry("620x430")
        window.resizable(True, True)
        window.minsize(520, 360)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="선택 항목 일괄 변경", style="Title.TLabel").pack(anchor="w")
        ttk.Label(body, text=f"선택된 항목: {len(items)}건", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 16))

        form = ttk.Frame(body)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="새 처리기한").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=7)
        deadline_var = tk.StringVar()
        ttk.Entry(form, textvariable=deadline_var).grid(row=0, column=1, sticky="ew", pady=7)
        ttk.Label(form, text=f"비워두면 변경하지 않습니다. ({DATE_HINT})", style="Subtitle.TLabel").grid(row=1, column=1, sticky="w")

        ttk.Label(form, text="진행 항목").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=7)
        progress_options = ["변경하지 않음", *[f"{index}개 항목 완료" for index in range(1, max_nodes + 1)]]
        progress_var = tk.StringVar(value=progress_options[0])
        ttk.Combobox(form, textvariable=progress_var, values=progress_options, state="readonly").grid(row=2, column=1, sticky="ew", pady=7)
        ttk.Label(form, text="각 업무에 설정된 단계 이름은 유지하고 같은 순번까지 완료합니다.", style="Subtitle.TLabel").grid(row=3, column=1, sticky="w")

        priority_var = tk.StringVar(value="변경하지 않음")
        if kind == "task":
            ttk.Label(form, text="비고").grid(row=4, column=0, sticky="w", padx=(0, 12), pady=7)
            ttk.Entry(form, textvariable=priority_var).grid(row=4, column=1, sticky="ew", pady=7)
            ttk.Label(form, text="비워두거나 '변경하지 않음'이면 변경하지 않습니다.", style="Subtitle.TLabel").grid(row=5, column=1, sticky="w")

        def save_bulk() -> None:
            deadline = deadline_var.get().strip()
            if deadline:
                try:
                    parse_deadline(deadline)
                except ValueError as exc:
                    messagebox.showwarning("날짜 확인", str(exc), parent=window)
                    return
            completed_count = 0 if progress_var.get() == "변경하지 않음" else int(progress_var.get().split("개", 1)[0])
            priority = "" if priority_var.get() in {"", "변경하지 않음"} else priority_var.get()
            if not deadline and not completed_count and not priority:
                messagebox.showinfo("변경할 내용", "변경할 항목을 하나 이상 입력하세요.", parent=window)
                return

            if kind == "task":
                self.db.bulk_update_tasks(item_ids, deadline, priority)
            for item_id, item in zip(item_ids, items):
                if kind == "complaint" and deadline and item["current_deadline"] != deadline:
                    self.db.extend_deadline(item_id, deadline, "선택 항목 일괄 변경")
                if completed_count:
                    if kind == "complaint":
                        stages = self._complaint_stages(item)
                        nodes = self._progress_node_keys(stages)
                        count = min(completed_count, len(nodes))
                        current = nodes[count] if count < len(nodes) else nodes[-1]
                        self.db.update_progress_state(item_id, {"current": current, "completed": nodes[:count]}, self._definition_text_from_key(current, stages))
                    else:
                        stages = self._task_stages(item)
                        nodes = self._progress_node_keys(stages)
                        count = min(completed_count, len(nodes))
                        current = nodes[count] if count < len(nodes) else nodes[-1]
                        self.db.update_task_progress_state(item_id, {"current": current, "completed": nodes[:count]}, self._task_stage_text(current, stages))
            self.refresh_all()
            window.destroy()

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(18, 0))
        ttk.Button(actions, text="취소", command=window.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(actions, text="일괄 변경 저장", style="Accent.TButton", command=save_bulk).pack(side="right")

    def show_audit_history(self, kind: str) -> None:
        tree = self.complaint_tree if kind == "complaint" else self.task_tree
        selected = tree.selection()
        if len(selected) != 1:
            messagebox.showinfo("항목 선택", "변경 이력을 확인할 항목 하나만 선택하세요.", parent=self)
            return
        item_id = int(selected[0])
        history = self.db.audit_history(kind, item_id)
        window = tk.Toplevel(self)
        window.title("변경 이력")
        window.transient(self)
        window.geometry("820x520")
        window.resizable(True, True)
        window.minsize(650, 400)
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="변경 이력", style="Title.TLabel").pack(anchor="w", pady=(0, 12))
        history_tree = self._new_tree(body, ("time", "action", "details"), ("변경 일시", "변경 내용", "세부 사항"), (180, 200, 400))
        for index, entry in enumerate(history):
            history_tree.insert("", "end", iid=str(index), values=(entry["created_at"], entry["action"], entry["details"]))

    def complete_complaint(self) -> None:
        item_ids = [int(iid) for iid in self.complaint_tree.selection()]
        if not item_ids:
            messagebox.showinfo("항목 선택", "완료할 민원을 선택하세요.", parent=self)
            return
        if messagebox.askyesno("처리 완료", f"선택한 민원 {len(item_ids)}건을 완료 처리할까요?", parent=self):
            for item_id in item_ids:
                self.db.mark_complete("complaint", item_id)
            self.refresh_all()

    def complete_task(self) -> None:
        item_ids = [int(iid) for iid in self.task_tree.selection()]
        if not item_ids:
            messagebox.showinfo("항목 선택", "완료할 수시업무를 선택하세요.", parent=self)
            return
        if messagebox.askyesno("처리 완료", f"선택한 수시업무 {len(item_ids)}건을 완료 처리할까요?", parent=self):
            for item_id in item_ids:
                self.db.mark_complete("task", item_id)
            self.refresh_all()

    def move_selected_to_trash(self, kind: str) -> None:
        tree = self.complaint_tree if kind == "complaint" else self.task_tree
        item_ids = [int(iid) for iid in tree.selection()]
        if not item_ids:
            messagebox.showinfo("항목 선택", "삭제할 항목을 먼저 선택하세요.", parent=self)
            return
        label = "민원" if kind == "complaint" else "수시업무"
        if not messagebox.askyesno("휴지통으로 이동", f"선택한 {label} {len(item_ids)}건을 휴지통으로 이동할까요?", parent=self):
            return
        self.db.move_to_trash(kind, item_ids)
        self.refresh_all()

    def restore_completed_item(self) -> None:
        selected = self.completed_tree.selection()
        if not selected:
            messagebox.showinfo("항목 선택", "복원할 완료 항목을 먼저 선택하세요.", parent=self)
            return
        if not messagebox.askyesno("완료 취소 및 복원", f"선택한 완료 항목 {len(selected)}건을 진행 상태로 복원할까요?", parent=self):
            return
        for iid in selected:
            kind, item_id = iid.split(":", 1)
            self.db.restore_completed(kind, int(item_id))
        self.refresh_all()

    def restore_trash_items(self) -> None:
        selected = self.trash_tree.selection()
        if not selected:
            messagebox.showinfo("항목 선택", "복원할 항목을 먼저 선택하세요.", parent=self)
            return
        if not messagebox.askyesno("휴지통 복원", f"선택한 항목 {len(selected)}건을 복원할까요?", parent=self):
            return
        grouped: dict[str, list[int]] = {"complaint": [], "task": []}
        for iid in selected:
            kind, item_id = iid.split(":", 1)
            grouped[kind].append(int(item_id))
        for kind, item_ids in grouped.items():
            self.db.restore_from_trash(kind, item_ids)
        self.refresh_all()

    def permanently_delete_trash(self) -> None:
        selected = self.trash_tree.selection()
        if not selected:
            messagebox.showinfo("항목 선택", "영구 삭제할 항목을 먼저 선택하세요.", parent=self)
            return
        if not messagebox.askyesno(
            "영구 삭제 확인",
            f"선택한 항목 {len(selected)}건을 영구 삭제할까요?\n이 작업은 되돌릴 수 없습니다.",
            icon="warning",
            parent=self,
        ):
            return
        grouped: dict[str, list[int]] = {"complaint": [], "task": []}
        for iid in selected:
            kind, item_id = iid.split(":", 1)
            grouped[kind].append(int(item_id))
        for kind, item_ids in grouped.items():
            self.db.permanently_delete(kind, item_ids)
        self.refresh_all()

    def open_task_excel_import(self) -> None:
        window = tk.Toplevel(self)
        window.title("수시업무 엑셀 가져오기")
        window.transient(self)
        window.geometry("620x230")
        window.resizable(True, True)
        window.minsize(520, 210)
        window.grab_set()

        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="수시업무 엑셀 가져오기", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text="샘플 양식을 내려받아 작성한 뒤 불러오면 여러 업무를 한 번에 등록할 수 있습니다.",
            style="Subtitle.TLabel",
            wraplength=560,
        ).pack(anchor="w", pady=(4, 16))

        actions = ttk.Frame(body)
        actions.pack(fill="x", pady=(22, 0))
        ttk.Button(
            actions,
            text="샘플 엑셀파일 다운로드",
            style="ExcelAction.TButton",
            command=lambda: self.download_task_excel_template(window),
        ).pack(side="left")
        ttk.Button(
            actions,
            text="작성한 엑셀 불러오기",
            style="ExcelAction.TButton",
            command=lambda: self.import_task_excel_file(window),
        ).pack(side="left", padx=8)
        ttk.Button(actions, text="닫기", style="ExcelAction.TButton", command=window.destroy).pack(side="right")

    def download_task_excel_template(self, parent: tk.Misc) -> None:
        selected = filedialog.asksaveasfilename(
            parent=parent,
            title="수시업무 샘플 엑셀 저장",
            defaultextension=".xlsx",
            initialfile="수시업무_가져오기_샘플.xlsx",
            filetypes=(("Excel 파일", "*.xlsx"),),
        )
        if not selected:
            return
        path = Path(selected)
        try:
            create_task_template(path)
        except Exception as exc:
            messagebox.showerror("샘플을 저장할 수 없음", str(exc), parent=parent)
            return
        messagebox.showinfo("샘플 저장 완료", f"샘플 엑셀파일을 저장했습니다.\n\n{path}", parent=parent)

    def import_task_excel_file(self, parent: tk.Misc) -> None:
        selected = filedialog.askopenfilename(
            parent=parent,
            title="수시업무 엑셀 선택",
            filetypes=(("Excel 파일", "*.xls *.xlsx *.xlsm"), ("모든 파일", "*.*")),
        )
        if not selected:
            return
        path = Path(selected)
        try:
            items = preview_task_import(path)
        except Exception as exc:
            messagebox.showerror("엑셀을 가져올 수 없음", str(exc), parent=parent)
            return
        if not items:
            messagebox.showinfo("등록할 업무 없음", "완전히 비어 있지 않은 업무 행을 찾지 못했습니다.", parent=parent)
            return
        if not messagebox.askyesno(
            "수시업무 일괄 등록",
            f"{path.name}에서 {len(items)}건을 수시업무로 등록할까요?\n빈 셀은 빈 값 그대로 등록됩니다.",
            parent=parent,
        ):
            return
        try:
            initial_node = self._progress_node_keys(self.task_stage_config)[0]
            initial_stage = self._task_stage_text(initial_node, self.task_stage_config)
            counts = apply_task_import(path, self.db, items, initial_stage)
        except Exception as exc:
            messagebox.showerror("수시업무를 등록할 수 없음", str(exc), parent=parent)
            return
        self.refresh_all()
        messagebox.showinfo("수시업무 등록 완료", f"수시업무 {counts['new']}건을 등록했습니다.", parent=parent)
        if parent.winfo_exists():
            parent.destroy()

    def open_excel_import(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self,
            title="받은민원 엑셀 선택",
            filetypes=(("Excel 파일", "*.xls *.xlsx *.xlsm"), ("모든 파일", "*.*")),
        )
        if not selected:
            return
        path = Path(selected)
        try:
            items = preview_import(path, self.db)
            counts = apply_import(path, self.db, items)
        except Exception as exc:
            messagebox.showerror("엑셀을 가져올 수 없음", str(exc), parent=self)
            return
        self.refresh_all()
        messagebox.showinfo(
            "엑셀 가져오기 완료",
            (
                f"신규 등록 {counts['new']}건\n"
                f"처리기한 갱신 {counts['changed']}건\n"
                f"기존과 동일 {counts['skipped']}건\n"
                f"오류 {counts['error']}건\n\n"
                "새올 민원 처리 화면에 자동 반영했습니다."
            ),
            parent=self,
        )

    def _open_excel_import_preview_legacy(self) -> None:
        selected = filedialog.askopenfilename(
            parent=self,
            title="받은민원 엑셀 선택",
            filetypes=(("Excel 파일", "*.xls *.xlsx *.xlsm"), ("모든 파일", "*.*")),
        )
        if not selected:
            return
        self.import_path = Path(selected)
        try:
            self.import_items = preview_import(self.import_path, self.db)
        except Exception as exc:
            messagebox.showerror("엑셀을 읽을 수 없음", str(exc), parent=self)
            return
        self.import_window = tk.Toplevel(self)
        self.import_window.title("민원 엑셀 가져오기")
        self.import_window.transient(self)
        self.import_window.geometry("1120x620")
        self.import_window.resizable(True, True)
        self.import_window.minsize(850, 480)
        self.import_window.grab_set()
        body = ttk.Frame(self.import_window, padding=18)
        body.pack(fill="both", expand=True)
        bar = self._toolbar(body)
        self.import_file_var = tk.StringVar(value="선택된 파일 없음")
        ttk.Label(bar, textvariable=self.import_file_var, style="Subtitle.TLabel").pack(side="left")
        self.import_button = ttk.Button(bar, text="확인한 내용 등록", state="disabled", style="Accent.TButton", command=self.commit_import)
        self.import_button.pack(side="right")
        ttk.Label(
            body,
            text="접수번호가 같으면 새 처리기한만 갱신하고, 기존 민원명·민원인·받은일자·처리단계는 변경하지 않습니다.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(0, 8))
        self.import_tree = self._new_tree(
            body,
            ("status", "row", "receipt", "name", "person", "received", "deadline", "message"),
            ("결과", "행", "접수번호", "민원명", "민원인", "받은일자", "처리기한", "확인사항"),
            (75, 55, 120, 200, 105, 105, 145, 180),
        )
        self.import_file_var.set(f"{self.import_path.name} · {len(self.import_items)}건")
        self._clear_tree(self.import_tree)
        for item in self.import_items:
            data = item.data
            tag = "error" if item.status == "오류" else "changed" if item.status == "변경" else "completed" if item.status == "중복" else "normal"
            self.import_tree.insert(
                "", "end", values=(item.status, item.row_number, data.get("receipt_no", ""),
                data.get("complaint_name", ""), data.get("applicant_name", ""), data.get("received_at", ""),
                data.get("deadline", ""), item.message), tags=(tag,)
            )
        can_import = any(item.status in {"신규", "변경"} for item in self.import_items)
        self.import_button.configure(state="normal" if can_import else "disabled")

    def commit_import(self) -> None:
        if not self.import_path or not self.import_items:
            return
        changed = sum(item.status == "변경" for item in self.import_items)
        prompt = "신규 건을 등록합니다."
        if changed:
            prompt += f"\n변경 {changed}건은 새 엑셀 파일의 처리기한으로 갱신합니다."
        if not messagebox.askyesno("엑셀 등록 확인", prompt + "\n계속할까요?", parent=self):
            return
        try:
            counts = apply_import(self.import_path, self.db, self.import_items)
        except Exception as exc:
            messagebox.showerror("등록 실패", str(exc), parent=self)
            return
        messagebox.showinfo(
            "엑셀 등록 완료",
            f"신규 {counts['new']}건\n변경 {counts['changed']}건\n중복 {counts['skipped']}건\n오류 {counts['error']}건",
            parent=self,
        )
        self.import_button.configure(state="disabled")
        self.refresh_all()
        if hasattr(self, "import_window") and self.import_window.winfo_exists():
            self.import_window.destroy()

    def _clear_tree(self, tree: ttk.Treeview) -> None:
        children = tree.get_children()
        if children:
            tree.delete(*children)
        self.after_idle(lambda: self._redraw_progress_if_needed(tree))

    def refresh_all(self) -> None:
        self.refresh_complaints()
        self.refresh_tasks()
        self.refresh_today()
        self.refresh_calendar()
        self.refresh_completed()
        self.refresh_trash()
        self.refresh_summary()

    def refresh_complaints(self) -> None:
        self._clear_tree(self.complaint_tree)
        self._complaint_progress.clear()
        query = self.complaint_search.get().strip().lower()
        complaints = self.db.list_complaints()
        for item in complaints:
            haystack = " ".join((item["receipt_no"], item["complaint_name"], item["applicant_name"], item["birth_date"])).lower()
            if query and query not in haystack:
                continue
            if not self._matches_quick_filter(item["current_deadline"], self.complaint_quick_filter.get()):
                continue
            self.complaint_tree.insert(
                "", "end", iid=str(item["id"]),
                values=(item["receipt_no"], item["complaint_name"], item["applicant_name"], item["received_at"],
                        item["current_deadline"], dday_text(item["current_deadline"]), "", ""),
                tags=(urgency_tag(item["current_deadline"]),),
            )
            self._complaint_progress[str(item["id"])] = (
                item, urgency_tag(item["current_deadline"])
            )
        self.after_idle(self._draw_complaint_progress)

    def refresh_tasks(self) -> None:
        self._clear_tree(self.task_tree)
        self._task_progress.clear()
        query = self.task_search.get().strip().lower()
        for item in self.db.list_tasks():
            haystack = " ".join((item["title"], item["content"])).lower()
            if query and query not in haystack:
                continue
            if not self._matches_quick_filter(item["deadline"], self.task_quick_filter.get(), item["priority"]):
                continue
            self.task_tree.insert(
                "", "end", iid=str(item["id"]),
                values=(item["title"], item["registered_at"], item["deadline"],
                        dday_text(item["deadline"]), item["priority"], "", ""), tags=(urgency_tag(item["deadline"]),),
            )
            self._task_progress[str(item["id"])] = (item, urgency_tag(item["deadline"]))
        self.after_idle(self._draw_task_progress)

    @staticmethod
    def _matches_quick_filter(deadline: str, mode: str, priority: str = "") -> bool:
        if mode == "전체":
            return True
        try:
            delta = (parse_deadline(deadline).date() - date.today()).days
        except ValueError:
            return False
        if mode == "오늘 마감":
            return delta == 0
        if mode == "3일 이내":
            return 0 <= delta <= 3
        if mode == "이번 주":
            return 0 <= delta <= 7
        if mode == "기한 초과":
            return delta < 0
        return True

    def _active_items(self) -> list[tuple[str, dict, str]]:
        result = [("민원 신청", item, item["current_deadline"]) for item in self.db.list_complaints()]
        result += [("수시업무", item, item["deadline"]) for item in self.db.list_tasks()]
        return result

    def _calendar_items(self) -> list[dict[str, object]]:
        items: list[dict[str, object]] = []
        for complaint in self.db.list_complaints():
            complaint_details = [f"진행 현황: {complaint['processing_stage']}"]
            if complaint.get("assignee"):
                complaint_details.append(f"담당자: {complaint['assignee']}")
            if complaint.get("memo"):
                complaint_details.append(f"메모: {complaint['memo']}")
            items.append(
                {
                    "kind": "complaint",
                    "label": "새올 민원",
                    "id": int(complaint["id"]),
                    "title": str(complaint["complaint_name"]),
                    "deadline": str(complaint["current_deadline"]),
                    "time": "",
                    "details": "\n".join(complaint_details),
                }
            )
        for task in self.db.list_tasks():
            task_details = [f"진행 현황: {task['processing_stage']}"]
            if task.get("priority"):
                task_details.append(f"비고: {task['priority']}")
            if task.get("content"):
                task_details.append(f"업무 내용: {task['content']}")
            items.append(
                {
                    "kind": "task",
                    "label": "수시업무",
                    "id": int(task["id"]),
                    "title": str(task["title"]),
                    "deadline": str(task["deadline"]),
                    "time": "",
                    "details": "\n".join(task_details),
                }
            )
        for event in self.db.list_calendar_events():
            event_time = str(event.get("event_time", ""))
            schedule_type = str(event.get("schedule_type", "개인일정"))
            items.append(
                {
                    "kind": "event",
                    "label": schedule_type,
                    "schedule_type": schedule_type,
                    "id": int(event["id"]),
                    "title": str(event["title"]),
                    "deadline": f"{event['event_date']} {event_time}".strip(),
                    "calendar_date": str(event["event_date"]),
                    "time": event_time,
                    "details": str(event["details"]),
                }
            )
        return items

    def _filtered_calendar_items(self) -> list[dict[str, object]]:
        mode = self.calendar_filter.get() if hasattr(self, "calendar_filter") else "전체"
        return [
            item
            for item in self._calendar_items()
            if mode == "전체"
            or (mode == "새올 민원" and item["kind"] == "complaint")
            or (mode == "수시업무" and item["kind"] == "task")
            or (mode == "개인일정" and item.get("schedule_type") == "개인일정")
            or (mode == "부서일정" and item.get("schedule_type") == "부서일정")
        ]

    @staticmethod
    def _calendar_deadline_date(item: dict[str, object]) -> date | None:
        try:
            return parse_deadline(str(item.get("calendar_date", item["deadline"]))).date()
        except ValueError:
            return None

    def change_calendar_month(self, direction: int) -> None:
        year = self.calendar_month.year
        month = self.calendar_month.month + direction
        if month < 1:
            year -= 1
            month = 12
        elif month > 12:
            year += 1
            month = 1
        self.calendar_month = date(year, month, 1)
        self.calendar_selected_date = self.calendar_month
        self.refresh_calendar()

    def go_calendar_today(self) -> None:
        today = date.today()
        self.calendar_month = today.replace(day=1)
        self.calendar_selected_date = today
        self.refresh_calendar()

    def select_calendar_date(self, selected: date) -> None:
        self.calendar_selected_date = selected
        self.refresh_calendar()

    def schedule_calendar_date_selection(self, selected: date) -> None:
        pending = getattr(self, "_calendar_click_job", None)
        if pending is not None:
            self.after_cancel(pending)

        def apply_selection() -> None:
            self._calendar_click_job = None
            self.select_calendar_date(selected)

        self._calendar_click_job = self.after(240, apply_selection)

    def add_calendar_event_dialog(self, selected: date, event: dict[str, object] | None = None) -> None:
        pending = getattr(self, "_calendar_click_job", None)
        if pending is not None:
            self.after_cancel(pending)
            self._calendar_click_job = None
        self.calendar_selected_date = selected
        editing = event is not None
        window = tk.Toplevel(self)
        window.title(f"일정 {'수정' if editing else '등록'} · {selected.isoformat()}")
        window.transient(self)
        window.geometry("620x520")
        window.minsize(520, 440)
        window.resizable(True, True)
        window.grab_set()
        body = ttk.Frame(window, padding=20)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(5, weight=1)

        action_label = "일정 수정" if editing else "일정 등록"
        ttk.Label(body, text=action_label, style="Title.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 14))
        ttk.Label(body, text="날짜 *").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=6)
        date_var = tk.StringVar(value=str(event.get("event_date", selected.isoformat())) if event else selected.isoformat())
        ttk.Entry(body, textvariable=date_var).grid(row=1, column=1, sticky="ew", pady=6)

        ttk.Label(body, text="일정 구분 *").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=6)
        schedule_type_var = tk.StringVar(
            value=str(event.get("schedule_type", "개인일정")) if event else "개인일정"
        )
        ttk.Combobox(
            body,
            textvariable=schedule_type_var,
            values=("개인일정", "부서일정"),
            state="readonly",
            width=18,
        ).grid(row=2, column=1, sticky="w", pady=6)

        ttk.Label(body, text="일정 제목 *").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=6)
        title_var = tk.StringVar(value=str(event.get("title", "")) if event else "")
        ttk.Entry(body, textvariable=title_var).grid(row=3, column=1, sticky="ew", pady=6)

        ttk.Label(body, text="시간 선택").grid(row=4, column=0, sticky="w", padx=(0, 12), pady=6)
        time_row = ttk.Frame(body)
        time_row.grid(row=4, column=1, sticky="w", pady=6)
        saved_time = str(event.get("event_time", "")) if event else ""
        saved_hour, saved_minute = (saved_time.split(":", 1) if ":" in saved_time else ("", ""))
        hour_var = tk.StringVar(value=f"{int(saved_hour)}시" if saved_hour else "")
        minute_var = tk.StringVar(value=f"{int(saved_minute)}분" if saved_minute else "")
        ttk.Combobox(
            time_row,
            textvariable=hour_var,
            values=("", *[f"{value}시" for value in range(25)]),
            state="readonly",
            width=8,
        ).pack(side="left")
        ttk.Combobox(
            time_row,
            textvariable=minute_var,
            values=("", *[f"{value}분" for value in range(61)]),
            state="readonly",
            width=8,
        ).pack(side="left", padx=(8, 0))
        ttk.Label(time_row, text="선택하지 않아도 등록됩니다.", style="Subtitle.TLabel").pack(side="left", padx=(10, 0))

        ttk.Label(body, text="상세 내용").grid(row=5, column=0, sticky="nw", padx=(0, 12), pady=6)
        details = tk.Text(body, height=8, wrap="word", font=self.input_font)
        details.insert("1.0", str(event.get("details", "")) if event else "")
        details.grid(row=5, column=1, sticky="nsew", pady=6)

        def save() -> None:
            title = title_var.get().strip()
            if not title:
                messagebox.showwarning("입력 확인", "일정 제목을 입력하세요.", parent=window)
                return
            try:
                event_date = parse_deadline(date_var.get().strip()).date()
            except ValueError as exc:
                messagebox.showwarning("날짜 확인", str(exc), parent=window)
                return
            hour_text = hour_var.get()
            minute_text = minute_var.get()
            if minute_text and not hour_text:
                messagebox.showwarning("시간 확인", "분을 선택하려면 시간도 함께 선택하세요.", parent=window)
                return
            event_time = ""
            if hour_text:
                hour = int(hour_text.removesuffix("시"))
                minute = int(minute_text.removesuffix("분")) if minute_text else 0
                event_time = f"{hour:02d}:{minute:02d}"
            if editing:
                self.db.update_calendar_event(
                    int(event["id"]),
                    event_date.isoformat(),
                    title,
                    details.get("1.0", "end").strip(),
                    event_time,
                    schedule_type_var.get(),
                )
            else:
                self.db.add_calendar_event(
                    event_date.isoformat(),
                    title,
                    details.get("1.0", "end").strip(),
                    event_time,
                    schedule_type_var.get(),
                )
            self.calendar_month = event_date.replace(day=1)
            self.calendar_selected_date = event_date
            self.refresh_all()
            window.destroy()

        actions = ttk.Frame(body)
        actions.grid(row=6, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(actions, text="취소", command=window.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(actions, text=action_label, style="Accent.TButton", command=save).pack(side="right")

    def refresh_calendar(self) -> None:
        if not hasattr(self, "calendar_days_frame"):
            return
        self.calendar_month_var.set(f"{self.calendar_month.year}년 {self.calendar_month.month}월")
        items = self._filtered_calendar_items()
        dated_items = [(item, self._calendar_deadline_date(item)) for item in items]

        for child in self.calendar_days_frame.winfo_children():
            child.destroy()
        weeks = calendar.Calendar(firstweekday=0).monthdatescalendar(
            self.calendar_month.year,
            self.calendar_month.month,
        )
        today = date.today()
        try:
            korean_holidays = holidays.country_holidays(
                "KR",
                years=[self.calendar_month.year],
                language="ko",
            )
        except (FileNotFoundError, OSError):
            # 배포 파일에 번역 리소스가 누락되더라도 달력 전체가 종료되지 않게 한다.
            korean_holidays = {}
        for row, week in enumerate(weeks):
            self.calendar_days_frame.rowconfigure(row, weight=1, uniform="calendar-row")
            for column, day in enumerate(week):
                in_month = day.month == self.calendar_month.month
                day_items = [item for item, item_date in dated_items if item_date == day] if in_month else []
                complaint_count = sum(item["kind"] == "complaint" for item in day_items)
                task_count = sum(item["kind"] == "task" for item in day_items)
                personal_event_count = sum(item.get("schedule_type") == "개인일정" for item in day_items)
                department_event_count = sum(item.get("schedule_type") == "부서일정" for item in day_items)

                selected = in_month and day == self.calendar_selected_date
                background = "#D8F0CD" if selected else "#FFFFFF" if in_month else "#EEF2EC"
                if not in_month:
                    foreground = "#A8B2A8"
                elif column == 5:
                    foreground = "#3478C5"
                elif column == 6:
                    foreground = "#D74343"
                else:
                    foreground = PALETTE["text"]
                border_color = (
                    PALETTE["green_dark"]
                    if selected
                    else PALETTE["green"]
                    if day == today and in_month
                    else "#D9E5D6"
                )
                day_canvas = tk.Canvas(
                    self.calendar_days_frame,
                    background=background,
                    highlightthickness=2 if selected or day == today and in_month else 1,
                    highlightbackground=border_color,
                    height=82,
                    cursor="hand2" if in_month else "",
                )
                day_canvas.grid(row=row, column=column, sticky="nsew", padx=1, pady=1)
                day_canvas.create_text(8, 7, anchor="nw", text=str(day.day), fill=foreground, font=(self.ui_font, 10, "bold"))
                holiday_name = str(korean_holidays.get(day, "")) if in_month else ""
                if holiday_name:
                    display_holiday = holiday_name if len(holiday_name) <= 15 else holiday_name[:14] + "…"
                    day_canvas.create_text(30, 9, anchor="nw", text=display_holiday, fill="#D74343", font=(self.ui_font, 8, "bold"))
                count_y = 30
                for count_label, count, color in (
                    ("새올", complaint_count, "#8A4DBF"),
                    ("수시업무", task_count, "#3478C5"),
                    ("개인", personal_event_count, "#202020"),
                    ("부서", department_event_count, "#155AA8"),
                ):
                    if count:
                        day_canvas.create_text(
                            8,
                            count_y,
                            anchor="nw",
                            text=f"{count_label} {count}",
                            fill=color,
                            font=(self.ui_font, 8, "bold"),
                        )
                        count_y += 13
                if in_month:
                    day_canvas.bind("<Button-1>", lambda _event, value=day: self.schedule_calendar_date_selection(value))
                    day_canvas.bind("<Double-1>", lambda _event, value=day: self.add_calendar_event_dialog(value))

        children = self.calendar_day_tree.get_children("")
        if children:
            self.calendar_day_tree.delete(*children)
        self._calendar_detail_items: dict[str, dict[str, object]] = {}

        selected_items = [
            item for item, item_date in dated_items
            if item_date == self.calendar_selected_date
        ]
        selected_items.sort(key=lambda item: (str(item["deadline"]), str(item["title"])))
        self.calendar_selected_var.set(
            f"{self.calendar_selected_date.year}년 {self.calendar_selected_date.month}월 "
            f"{self.calendar_selected_date.day}일 · {len(selected_items)}건"
        )
        for item in selected_items:
            iid = f"{item['kind']}:{item['id']}"
            self.calendar_day_tree.insert(
                "",
                "end",
                iid=iid,
                values=(item["label"], item["title"], item["time"]),
                tags=(("department_event" if item.get("schedule_type") == "부서일정" else "personal_event")
                      if item["kind"] == "event" else str(item["kind"]),),
            )
            self._calendar_detail_items[iid] = item
        rows = self.calendar_day_tree.get_children("")
        if rows:
            self.calendar_day_tree.selection_set(rows[0])
            self.calendar_day_tree.focus(rows[0])
        self.show_calendar_detail()

    def show_calendar_detail(self) -> None:
        if not hasattr(self, "calendar_detail_var"):
            return
        selected = self.calendar_day_tree.selection()
        item = self._calendar_detail_items.get(selected[0]) if selected else None
        if not item:
            self.calendar_detail_title_var.set("선택한 날짜에 표시할 항목이 없습니다.")
            self.calendar_detail_var.set("")
            return
        self.calendar_detail_title_var.set(f"[{item['label']}] {item['title']}")
        details = str(item.get("details", "")).strip()
        self.calendar_detail_var.set(details or "입력된 상세 내용이 없습니다.")

    def _selected_calendar_item(self) -> dict[str, object] | None:
        selected = self.calendar_day_tree.selection()
        if not selected:
            messagebox.showinfo("항목 선택", "수정하거나 삭제할 항목을 먼저 선택하세요.", parent=self)
            return None
        return self._calendar_detail_items.get(selected[0])

    def edit_selected_calendar_item(self) -> None:
        item = self._selected_calendar_item()
        if not item:
            return
        kind = str(item["kind"])
        item_id = int(item["id"])
        if kind == "event":
            event = self.db.calendar_event_by_id(item_id)
            if event:
                self.add_calendar_event_dialog(
                    parse_deadline(str(event["event_date"])).date(),
                    event,
                )
            return
        if kind == "complaint":
            self.show_page(1)
            self.complaint_tree.selection_set(str(item_id))
            self.complaint_tree.see(str(item_id))
            self.edit_complaint_dialog()
        elif kind == "task":
            self.show_page(2)
            self.task_tree.selection_set(str(item_id))
            self.task_tree.see(str(item_id))
            self.edit_task_dialog()

    def delete_selected_calendar_item(self) -> None:
        item = self._selected_calendar_item()
        if not item:
            return
        kind = str(item["kind"])
        item_id = int(item["id"])
        label = str(item["label"])
        title = str(item["title"])
        action_text = "삭제" if kind == "event" else "휴지통으로 이동"
        if not messagebox.askyesno(
            "선택 항목 삭제",
            f"[{label}] {title}\n\n이 항목을 {action_text}할까요?",
            parent=self,
        ):
            return
        if kind == "event":
            self.db.delete_calendar_event(item_id)
        else:
            self.db.move_to_trash(kind, [item_id])
        self.refresh_all()

    def open_calendar_item(self, tree: ttk.Treeview, event: tk.Event | None = None) -> None:
        row_id = tree.identify_row(event.y) if event is not None else ""
        if not row_id:
            selected = tree.selection()
            row_id = selected[0] if selected else ""
        if not row_id or ":" not in row_id:
            return
        kind, item_id = row_id.split(":", 1)
        self._show_source_item(kind, item_id)

    def refresh_today(self) -> None:
        self._clear_tree(self.today_tree)
        for kind, item, deadline in self._active_items():
            try:
                delta = (parse_deadline(deadline).date() - date.today()).days
            except ValueError:
                delta = 9999
            if delta > 7:
                continue
            is_complaint = kind == "민원 신청"
            self.today_tree.insert(
                "", "end", iid=f"{'complaint' if is_complaint else 'task'}:{item['id']}",
                values=(kind, item["complaint_name"] if is_complaint else item["title"], item["applicant_name"] if is_complaint else "",
                        deadline, dday_text(deadline), item["processing_stage"] if is_complaint else item["priority"]),
                tags=(urgency_tag(deadline),),
            )
        today_text = date.today().isoformat()
        today_events = [
            event for event in self.db.list_calendar_events()
            if str(event["event_date"]) == today_text
        ]
        today_events.sort(
            key=lambda event: (
                not bool(str(event.get("event_time", ""))),
                str(event.get("event_time", "")),
                int(event["id"]),
            )
        )
        for event in today_events:
            event_time = str(event.get("event_time", ""))
            date_and_time = f"{today_text} {event_time}".strip()
            self.today_tree.insert(
                "",
                "end",
                iid=f"event:{event['id']}",
                values=(event.get("schedule_type", "개인일정"), event["title"], "", date_and_time, "오늘", event.get("schedule_type", "개인일정")),
                tags=(("department_event" if event.get("schedule_type") == "부서일정" else "personal_event"),),
            )

    def open_today_item(self, event: tk.Event | None = None) -> None:
        if event is not None:
            row_id = self.today_tree.identify_row(event.y)
        else:
            selected = self.today_tree.selection()
            row_id = selected[0] if selected else ""
        if not row_id or ":" not in row_id:
            return
        kind, item_id = row_id.split(":", 1)
        self._show_source_item(kind, item_id)

    def _show_source_item(self, kind: str, item_id: str) -> None:
        if kind == "complaint":
            self.complaint_search.set("")
            self.complaint_quick_filter.set("전체")
            self.refresh_complaints()
            target_tree = self.complaint_tree
            page_index = 1
        elif kind == "task":
            self.task_search.set("")
            self.task_quick_filter.set("전체")
            self.refresh_tasks()
            target_tree = self.task_tree
            page_index = 2
        elif kind == "event":
            event = self.db.calendar_event_by_id(int(item_id))
            if not event:
                return
            event_date = parse_deadline(str(event["event_date"])).date()
            self.calendar_month = event_date.replace(day=1)
            self.calendar_selected_date = event_date
            self.calendar_filter.set("전체")
            self.show_page(3)
            self.refresh_calendar()
            calendar_iid = f"event:{item_id}"
            if self.calendar_day_tree.exists(calendar_iid):
                self.calendar_day_tree.selection_set(calendar_iid)
                self.calendar_day_tree.focus(calendar_iid)
                self.calendar_day_tree.see(calendar_iid)
                self.show_calendar_detail()
            return
        else:
            return
        self.show_page(page_index)
        if target_tree.exists(item_id):
            target_tree.selection_set(item_id)
            target_tree.focus(item_id)
            target_tree.see(item_id)

    def refresh_completed(self) -> None:
        self._clear_tree(self.completed_tree)
        for item in self.db.list_complaints(include_completed=True):
            if item["status"] == "완료":
                self.completed_tree.insert("", "end", iid=f"complaint:{item['id']}", values=("민원 신청", item["complaint_name"], item["applicant_name"], item["current_deadline"], item["completed_at"]), tags=("completed",))
        for item in self.db.list_tasks(include_completed=True):
            if item["status"] == "완료":
                self.completed_tree.insert("", "end", iid=f"task:{item['id']}", values=("수시업무", item["title"], "", item["deadline"], item["completed_at"]), tags=("completed",))

    def refresh_trash(self) -> None:
        self._clear_tree(self.trash_tree)
        for item in self.db.list_complaints(include_completed=True):
            if item["status"] == "삭제":
                self.trash_tree.insert(
                    "", "end", iid=f"complaint:{item['id']}",
                    values=("민원 신청", item["complaint_name"], item["applicant_name"], item["current_deadline"], item["updated_at"]),
                    tags=("completed",),
                )
        for item in self.db.list_tasks(include_completed=True):
            if item["status"] == "삭제":
                self.trash_tree.insert(
                    "", "end", iid=f"task:{item['id']}",
                    values=("수시업무", item["title"], "", item["deadline"], item["updated_at"]),
                    tags=("completed",),
                )

    def refresh_summary(self) -> None:
        complaints = self.db.list_complaints()
        tasks = self.db.list_tasks()
        deadlines = [item["current_deadline"] for item in complaints] + [item["deadline"] for item in tasks]
        deltas = []
        for value in deadlines:
            try:
                deltas.append((parse_deadline(value).date() - date.today()).days)
            except ValueError:
                pass
        self.summary_vars["민원"].set(str(len(complaints)))
        self.summary_vars["수시업무"].set(str(len(tasks)))
        today_text = date.today().isoformat()
        today_event_count = sum(
            str(event["event_date"]) == today_text
            for event in self.db.list_calendar_events()
        )
        self.summary_vars["오늘의 일정"].set(str(today_event_count))
        self.summary_vars["오늘 마감"].set(str(sum(delta == 0 for delta in deltas)))
        self.summary_vars["3일 이내"].set(str(sum(0 < delta <= 3 for delta in deltas)))
        self.summary_vars["기한 초과"].set(str(sum(delta < 0 for delta in deltas)))

    def _notify_due_items(self) -> None:
        today_count = int(self.summary_vars["오늘 마감"].get())
        overdue_count = int(self.summary_vars["기한 초과"].get())
        if today_count or overdue_count:
            show_windows_notification(
                APP_TITLE,
                f"오늘 마감 {today_count}건 · 기한 초과 {overdue_count}건이 있습니다. 앱에서 확인하세요.",
            )

    def show_notification_sample(self) -> None:
        shown = show_windows_notification(
            f"{APP_TITLE} - 알림 샘플",
            "오늘 마감 2건 · 3일 이내 4건 · 기한 초과 1건이 있습니다. 앱에서 확인하세요.",
        )
        if not shown:
            messagebox.showwarning("알림 샘플", "Windows 알림을 표시하지 못했습니다.", parent=self)

    def on_close(self) -> None:
        self._save_stage_config()
        self._save_task_stage_config()
        self.db.close()
        self.destroy()


def unlock_database(db: Database) -> bool:
    helper = tk.Tk()
    helper.withdraw()
    try:
        if db.try_auto_unlock():
            helper.destroy()
            return True
    except (OSError, InvalidPasswordError, ValueError):
        pass
    if not db.has_password():
        password = secrets.token_urlsafe(32)
        db.set_initial_password(password)
        db.enable_auto_unlock(password)
        helper.destroy()
        return True
    messagebox.showinfo(
        "자동 시작 전환",
        "앞으로 비밀번호를 묻지 않도록 전환합니다.\n기존 데이터 확인을 위해 현재 비밀번호를 이번 한 번만 입력하세요.",
        parent=helper,
    )
    for _ in range(5):
        password = simpledialog.askstring(APP_TITLE, "기존 비밀번호", show="*", parent=helper)
        if password is None:
            helper.destroy()
            return False
        try:
            db.unlock(password)
            db.enable_auto_unlock(password)
            helper.destroy()
            return True
        except InvalidPasswordError:
            messagebox.showerror("로그인 실패", "비밀번호가 올바르지 않습니다.", parent=helper)
    helper.destroy()
    return False


def main() -> int:
    db = Database()
    if not unlock_database(db):
        db.close()
        return 1
    app = ScheduleApp(db)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
