"""PyQt6 GUI.

新版界面目标:
* Excel 与 PPT 功能分开
* Material 风格（Google UI 风格参考）
* 用“INAND Benchmark”字符生成应用图标
"""

from __future__ import annotations

import logging
import subprocess
import sys
import traceback
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from PyQt6.QtCore import QObject, QThread, Qt, pyqtSignal
from PyQt6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDateEdit,
    QDoubleSpinBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .cli import run as cli_run
from .data_loader import resolve_data_source

logger = logging.getLogger("benchmark.gui")


def _resource_path(relative: str) -> Path:
    """返回开发态/打包态都可用的资源路径."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return base / relative


class _LogStream(logging.Handler, QObject):
    new_message = pyqtSignal(str)

    def __init__(self) -> None:
        logging.Handler.__init__(self)
        QObject.__init__(self)
        self.setFormatter(
            logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")
        )

    def emit(self, record: logging.LogRecord) -> None:  # type: ignore[override]
        try:
            self.new_message.emit(self.format(record))
        except Exception:
            self.handleError(record)


class _TaskWorker(QObject):
    finished = pyqtSignal(dict)
    failed = pyqtSignal(str)

    def __init__(self, task_name: str, runner: Callable[[], dict]):
        super().__init__()
        self.task_name = task_name
        self.runner = runner

    def run(self) -> None:
        try:
            out = self.runner()
            out["_task_name"] = self.task_name
            self.finished.emit(out)
        except Exception as exc:
            logger.exception("任务失败(%s): %s", self.task_name, exc)
            self.failed.emit(f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}")


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("INAND Benchmark Studio")
        self.setWindowIcon(self._build_text_icon("INAND\nBenchmark"))
        self.resize(1080, 780)

        self._thread: QThread | None = None
        self._worker: _TaskWorker | None = None

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)

        self._apply_material_style()
        root.addLayout(self._build_header())
        root.addWidget(self._build_excel_card())
        root.addWidget(self._build_ppt_card())
        root.addWidget(self._build_recent_card())
        root.addWidget(self._build_log_card(), stretch=1)

        self._stream = _LogStream()
        self._stream.new_message.connect(self.log_view.append)
        logging.getLogger().addHandler(self._stream)
        logging.getLogger().setLevel(logging.INFO)
        logger.info("GUI ready.")

    def _build_header(self) -> QHBoxLayout:
        box = QHBoxLayout()
        icon = QLabel("INAND\nBenchmark")
        icon.setObjectName("brandBadge")
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setFixedSize(110, 56)
        box.addWidget(icon)

        tbox = QVBoxLayout()
        title = QLabel("INAND Benchmark Studio")
        title.setObjectName("mainTitle")
        sub = QLabel("Excel 生成与 PPT 更新分离操作")
        sub.setObjectName("subtitle")
        tbox.addWidget(title)
        tbox.addWidget(sub)
        box.addLayout(tbox, stretch=1)

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        self.progress.setFixedWidth(220)
        box.addWidget(self.progress)
        return box

    def _build_excel_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        lay = QGridLayout(card)
        lay.setHorizontalSpacing(10)
        lay.setVerticalSpacing(8)

        head = QLabel("Excel 生成")
        head.setObjectName("cardTitle")
        lay.addWidget(head, 0, 0, 1, 2)

        self.data_edit = self._make_path_input("PN_TEST_STEPID_YIELD.txt")
        self.rule_edit = self._make_path_input("Rule_list.xlsx")
        self.src_edit = self._make_path_input("bachmark SRC.xlsx")
        self.output_edit = self._make_path_input("output/INAND_weekly_benchmark.xlsx")

        lay.addWidget(QLabel("原始数据 (.txt/.csv)"), 1, 0)
        lay.addLayout(self._with_browse(self.data_edit, "open", "*.txt *.csv"), 1, 1)
        lay.addWidget(QLabel("规则文件 (.xlsx)"), 2, 0)
        lay.addLayout(self._with_browse(self.rule_edit, "open", "*.xlsx"), 2, 1)
        lay.addWidget(QLabel("Goal SRC (.xlsx)"), 3, 0)
        lay.addLayout(self._with_browse(self.src_edit, "open", "*.xlsx"), 3, 1)
        lay.addWidget(QLabel("输出 Excel"), 4, 0)
        lay.addLayout(self._with_browse(self.output_edit, "save", "*.xlsx"), 4, 1)

        self.use_today_cb = QCheckBox("使用电脑当前日期（实时）")
        self.use_today_cb.setChecked(True)
        self.use_today_cb.toggled.connect(self._toggle_date)
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(date.today())
        self.date_edit.setEnabled(False)
        self.weeks_spin = QSpinBox()
        self.weeks_spin.setRange(1, 26)
        self.weeks_spin.setValue(9)

        date_row = QHBoxLayout()
        date_row.addWidget(self.use_today_cb)
        date_row.addWidget(self.date_edit)
        date_row.addStretch(1)
        date_row.addWidget(QLabel("周窗口"))
        date_row.addWidget(self.weeks_spin)
        lay.addLayout(date_row, 5, 0, 1, 2)

        self.preview_label = QLabel("推算预览: 运行后显示 current/prev/week-window/month/quarter")
        self.preview_label.setObjectName("hint")
        lay.addWidget(self.preview_label, 6, 0, 1, 2)

        btns = QHBoxLayout()
        self.run_excel_btn = QPushButton("生成 Excel")
        self.run_excel_btn.setObjectName("primaryBtn")
        self.run_excel_btn.clicked.connect(self.on_run_excel)
        self.open_output_btn = QPushButton("打开输出目录")
        self.open_output_btn.clicked.connect(self.on_open_output)
        btns.addWidget(self.run_excel_btn)
        btns.addWidget(self.open_output_btn)
        btns.addStretch(1)
        lay.addLayout(btns, 7, 0, 1, 2)
        return card

    def _build_ppt_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        lay = QGridLayout(card)
        lay.setHorizontalSpacing(10)
        lay.setVerticalSpacing(8)

        head = QLabel("PPT 更新")
        head.setObjectName("cardTitle")
        lay.addWidget(head, 0, 0, 1, 2)

        self.report_edit = self._make_path_input("output/INAND_weekly_benchmark.xlsx")
        self.ppt_edit = self._make_path_input("SDSS INAND YIELD WW45_2026_benchmark.pptx")
        self.rule_for_ppt_edit = self._make_path_input("Rule_list.xlsx")

        lay.addWidget(QLabel("用于粘贴的 Excel"), 1, 0)
        lay.addLayout(self._with_browse(self.report_edit, "open", "*.xlsx"), 1, 1)
        lay.addWidget(QLabel("目标 PPT"), 2, 0)
        lay.addLayout(self._with_browse(self.ppt_edit, "open", "*.pptx"), 2, 1)
        lay.addWidget(QLabel("顺序规则文件"), 3, 0)
        lay.addLayout(self._with_browse(self.rule_for_ppt_edit, "open", "*.xlsx"), 3, 1)

        self.left_spin = QDoubleSpinBox()
        self.left_spin.setDecimals(2)
        self.left_spin.setRange(0.0, 5.0)
        self.left_spin.setValue(0.2)
        self.top_spin = QDoubleSpinBox()
        self.top_spin.setDecimals(2)
        self.top_spin.setRange(0.0, 5.0)
        self.top_spin.setValue(1.5)
        pos_row = QHBoxLayout()
        pos_row.addWidget(QLabel("Left(inch)"))
        pos_row.addWidget(self.left_spin)
        pos_row.addWidget(QLabel("Top(inch)"))
        pos_row.addWidget(self.top_spin)
        pos_row.addStretch(1)
        lay.addLayout(pos_row, 4, 0, 1, 2)

        hint = QLabel("会自动删除目标页旧表格，再贴入最新表格。")
        hint.setObjectName("hint")
        lay.addWidget(hint, 5, 0, 1, 2)

        btns = QHBoxLayout()
        self.run_ppt_btn = QPushButton("更新 PPT")
        self.run_ppt_btn.setObjectName("primaryBtn")
        self.run_ppt_btn.clicked.connect(self.on_run_ppt)
        self.run_all_btn = QPushButton("一键 Excel + PPT")
        self.run_all_btn.clicked.connect(self.on_run_all)
        btns.addWidget(self.run_ppt_btn)
        btns.addWidget(self.run_all_btn)
        btns.addStretch(1)
        lay.addLayout(btns, 6, 0, 1, 2)
        return card

    def _build_log_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(12, 12, 12, 12)
        lay.addWidget(QLabel("运行日志"))
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFont(QFont("Consolas", 9))
        self.log_view.setStyleSheet(
            "QTextEdit { background:#0f172a; color:#e2e8f0; border-radius:8px; padding:8px; }"
        )
        lay.addWidget(self.log_view, stretch=1)
        return card

    def _build_recent_card(self) -> QFrame:
        card = QFrame()
        card.setObjectName("card")
        lay = QGridLayout(card)
        lay.setHorizontalSpacing(10)
        lay.setVerticalSpacing(6)
        title = QLabel("最近一次运行结果")
        title.setObjectName("cardTitle")
        lay.addWidget(title, 0, 0, 1, 4)

        lay.addWidget(QLabel("任务"), 1, 0)
        self.last_task_val = QLabel("-")
        self.last_task_val.setObjectName("hint")
        lay.addWidget(self.last_task_val, 1, 1)

        lay.addWidget(QLabel("时间"), 1, 2)
        self.last_time_val = QLabel("-")
        self.last_time_val.setObjectName("hint")
        lay.addWidget(self.last_time_val, 1, 3)

        lay.addWidget(QLabel("Excel"), 2, 0)
        self.last_excel_val = QLabel("-")
        self.last_excel_val.setObjectName("hint")
        self.last_excel_val.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        lay.addWidget(self.last_excel_val, 2, 1, 1, 3)

        lay.addWidget(QLabel("PPT"), 3, 0)
        self.last_ppt_val = QLabel("-")
        self.last_ppt_val.setObjectName("hint")
        self.last_ppt_val.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        lay.addWidget(self.last_ppt_val, 3, 1, 1, 3)

        lay.addWidget(QLabel("状态"), 4, 0)
        self.last_status_val = QLabel("未运行")
        self.last_status_val.setObjectName("hint")
        lay.addWidget(self.last_status_val, 4, 1, 1, 3)
        return card

    def _apply_material_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget { background:#f5f7fb; color:#1f2937; font-family:'Segoe UI'; font-size:13px; }
            #brandBadge { background:#1a73e8; color:white; border-radius:10px; font-weight:700; }
            #mainTitle { font-size:22px; font-weight:700; color:#111827; }
            #subtitle { color:#6b7280; }
            #card { background:white; border:1px solid #e5e7eb; border-radius:12px; }
            #cardTitle { font-size:16px; font-weight:700; color:#1f2937; padding:2px 0 6px 0; }
            #hint { color:#475569; }
            QLineEdit, QDateEdit, QSpinBox, QDoubleSpinBox {
                background:white; border:1px solid #d1d5db; border-radius:8px; padding:6px 8px;
            }
            QLineEdit:focus, QDateEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus { border:1px solid #1a73e8; }
            QPushButton {
                background:#e5e7eb; border:0; border-radius:8px; padding:8px 14px; font-weight:600;
            }
            QPushButton:hover { background:#dbe0e8; }
            #primaryBtn { background:#1a73e8; color:white; }
            #primaryBtn:hover { background:#1669c1; }
            """
        )

    def _build_text_icon(self, text: str) -> QIcon:
        pix = QPixmap(96, 96)
        pix.fill(QColor("#1a73e8"))
        p = QPainter(pix)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(QColor("#ffffff"))
        f = QFont("Segoe UI", 13, QFont.Weight.Bold)
        p.setFont(f)
        p.drawText(pix.rect(), Qt.AlignmentFlag.AlignCenter, text)
        p.end()
        return QIcon(pix)

    def _make_path_input(self, default: str) -> QLineEdit:
        e = QLineEdit(default)
        e.setMinimumHeight(30)
        return e

    def _with_browse(self, edit: QLineEdit, mode: str, filter_: str) -> QHBoxLayout:
        box = QHBoxLayout()
        box.addWidget(edit)
        btn = QPushButton("...")
        btn.setFixedWidth(42)
        btn.clicked.connect(lambda: self._pick_path(edit, mode, filter_))
        box.addWidget(btn)
        return box

    def _pick_path(self, edit: QLineEdit, mode: str, filter_: str) -> None:
        if mode == "open":
            p, _ = QFileDialog.getOpenFileName(self, "选择文件", edit.text(), filter_)
        else:
            p, _ = QFileDialog.getSaveFileName(self, "保存为", edit.text(), filter_)
        if p:
            edit.setText(p)

    def _toggle_date(self, checked: bool) -> None:
        self.date_edit.setEnabled(not checked)

    def _set_busy(self, busy: bool) -> None:
        self.run_excel_btn.setEnabled(not busy)
        self.run_ppt_btn.setEnabled(not busy)
        self.run_all_btn.setEnabled(not busy)
        self.progress.setVisible(busy)

    def _run_task(self, task_name: str, runner: Callable[[], dict]) -> None:
        if self._thread is not None and self._thread.isRunning():
            self._error("当前已有任务在运行，请稍后")
            return
        self._set_busy(True)
        self.log_view.clear()
        logger.info("开始任务: %s", task_name)
        self.last_task_val.setText(task_name)
        self.last_time_val.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.last_status_val.setText("运行中…")
        self._thread = QThread(self)
        self._worker = _TaskWorker(task_name, runner)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.finished.connect(self._on_task_finished)
        self._worker.failed.connect(self._on_task_failed)
        self._worker.finished.connect(self._thread.quit)
        self._worker.failed.connect(self._thread.quit)
        self._thread.finished.connect(self._thread.deleteLater)
        self._thread.start()

    def _collect_excel_params(self) -> tuple[Path, Path, Path, Path | None, datetime | None, int]:
        requested_data = Path(self.data_edit.text().strip())
        rule_path = Path(self.rule_edit.text().strip())
        output_path = Path(self.output_edit.text().strip())

        data_path = resolve_data_source(requested_data)
        if data_path != requested_data:
            logger.info("原始数据自动优先选择: %s (from %s)", data_path, requested_data)
        if not rule_path.exists():
            raise FileNotFoundError(f"找不到规则文件: {rule_path}")

        src_text = self.src_edit.text().strip()
        src_path: Path | None = Path(src_text) if src_text else None
        if src_path is not None and not src_path.exists():
            logger.warning("SRC 文件 %s 不存在，回退到原始数据目标", src_path)
            src_path = None

        today = None if self.use_today_cb.isChecked() else datetime(
            self.date_edit.date().year(),
            self.date_edit.date().month(),
            self.date_edit.date().day(),
        )
        return data_path, rule_path, output_path, src_path, today, self.weeks_spin.value()

    def on_run_excel(self) -> None:
        try:
            data_path, rule_path, output_path, src_path, today, weeks = self._collect_excel_params()
        except Exception as exc:
            self._error(str(exc))
            return

        def runner() -> dict:
            meta = cli_run(
                data_path=data_path,
                rule_path=rule_path,
                output_path=output_path,
                src_goals_path=src_path,
                today=today,
                weeks_count=weeks,
                log_file=Path("logs/gui.log"),
            )
            return meta

        self._run_task("excel", runner)

    def on_run_ppt(self) -> None:
        report_path = Path(self.report_edit.text().strip())
        ppt_path = Path(self.ppt_edit.text().strip())
        rule_path = Path(self.rule_for_ppt_edit.text().strip())
        if not report_path.exists():
            self._error(f"找不到 Excel 文件: {report_path}")
            return
        if not ppt_path.exists():
            self._error(f"找不到 PPT 文件: {ppt_path}")
            return
        if not rule_path.exists():
            self._error(f"找不到规则文件: {rule_path}")
            return
        left = self.left_spin.value()
        top = self.top_spin.value()

        def runner() -> dict:
            script = _resource_path("scripts/paste_excel_to_ppt.ps1")
            cmd = [
                "powershell",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-RulePath",
                str(rule_path),
                "-ReportPath",
                str(report_path),
                "-PptPath",
                str(ppt_path),
                "-LeftInch",
                str(left),
                "-TopInch",
                str(top),
            ]
            proc = subprocess.run(cmd, capture_output=True, text=True, cwd=str(Path.cwd()))
            out_lines = [x for x in (proc.stdout or "").splitlines() if x.strip()]
            err_lines = [x for x in (proc.stderr or "").splitlines() if x.strip()]
            for line in out_lines:
                logger.info(line)
            for line in err_lines:
                logger.warning(line)
            if proc.returncode != 0:
                raise RuntimeError("PPT 更新失败，请看日志")
            return {"ppt": str(ppt_path), "excel": str(report_path)}

        self._run_task("ppt", runner)

    def on_run_all(self) -> None:
        try:
            data_path, rule_path, output_path, src_path, today, weeks = self._collect_excel_params()
        except Exception as exc:
            self._error(str(exc))
            return
        ppt_path = Path(self.ppt_edit.text().strip())
        rule_for_ppt = Path(self.rule_for_ppt_edit.text().strip())
        left = self.left_spin.value()
        top = self.top_spin.value()

        def runner() -> dict:
            meta = cli_run(
                data_path=data_path,
                rule_path=rule_path,
                output_path=output_path,
                src_goals_path=src_path,
                today=today,
                weeks_count=weeks,
                log_file=Path("logs/gui.log"),
            )
            excel_path = str(meta.get("output") or meta.get("output_actual") or output_path)
            logger.info("Excel generated: %s", excel_path)
            script = _resource_path("scripts/paste_excel_to_ppt.ps1")
            cmd = [
                "powershell",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-RulePath",
                str(rule_for_ppt),
                "-ReportPath",
                excel_path,
                "-PptPath",
                str(ppt_path),
                "-LeftInch",
                str(left),
                "-TopInch",
                str(top),
            ]
            proc = subprocess.run(cmd, capture_output=True, text=True, cwd=str(Path.cwd()))
            out_lines = [x for x in (proc.stdout or "").splitlines() if x.strip()]
            err_lines = [x for x in (proc.stderr or "").splitlines() if x.strip()]
            for line in out_lines:
                logger.info(line)
            for line in err_lines:
                logger.warning(line)
            if proc.returncode != 0:
                raise RuntimeError("一键任务失败，请看日志")
            return {"excel": excel_path, "ppt": str(ppt_path)}

        self._run_task("all", runner)

    def _on_task_finished(self, meta: dict) -> None:
        self._set_busy(False)
        task = meta.get("_task_name", "?")
        logger.info("任务完成: %s", task)
        self.last_task_val.setText(task)
        self.last_time_val.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.last_status_val.setText("成功")

        if task == "excel":
            per = meta.get("periods", {})
            weeks = per.get("weeks_window", []) or []
            weeks_disp = f"{weeks[0]}..{weeks[-1]} ({len(weeks)} weeks)" if weeks else "<none>"
            self.preview_label.setText(
                f"推算结果: current={per.get('current_week')} prev={per.get('prev_week')} "
                f"weeks={weeks_disp} month={per.get('month')} quarter={per.get('quarter')}"
            )
            out = meta.get("output") or meta.get("output_actual")
            if out:
                self.output_edit.setText(str(out))
                self.report_edit.setText(str(out))
                self.last_excel_val.setText(str(out))
            QMessageBox.information(self, "完成", f"Excel 生成完成\n\n输出: {out}")
            return

        if task == "ppt":
            self.last_ppt_val.setText(str(meta.get("ppt", "-")))
            self.last_excel_val.setText(str(meta.get("excel", self.last_excel_val.text())))
            QMessageBox.information(self, "完成", f"PPT 已更新\n\n文件: {meta.get('ppt')}")
            return

        if task == "all":
            if meta.get("excel"):
                self.last_excel_val.setText(str(meta.get("excel")))
                self.output_edit.setText(str(meta.get("excel")))
                self.report_edit.setText(str(meta.get("excel")))
            if meta.get("ppt"):
                self.last_ppt_val.setText(str(meta.get("ppt")))
            QMessageBox.information(
                self,
                "完成",
                f"一键任务完成\n\nExcel: {meta.get('excel')}\nPPT: {meta.get('ppt')}",
            )
            return

    def _on_task_failed(self, msg: str) -> None:
        self._set_busy(False)
        logger.error(msg)
        self.last_time_val.setText(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.last_status_val.setText("失败")
        self._error(msg)

    def _error(self, msg: str) -> None:
        QMessageBox.critical(self, "错误", msg)

    def on_open_output(self) -> None:
        target = Path(self.output_edit.text()).parent.resolve()
        target.mkdir(parents=True, exist_ok=True)
        try:
            import os
            os.startfile(str(target))  # type: ignore[attr-defined]
        except AttributeError:
            subprocess.Popen(["xdg-open", str(target)])


def main() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
