# -*- coding: utf-8 -*-

import os, sys, json, re, subprocess
from typing import List, Dict, Any, Tuple, Optional
import fitz  # PyMuPDF
from PyQt5.QtGui import QIcon

from PyQt5.QtCore import Qt, QPoint, QUrl
from PyQt5.QtGui import QDesktopServices
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QGridLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QTextEdit,
    QMessageBox, QComboBox, QTabWidget, QMenu, QStyledItemDelegate
)

import sys, os
from PyQt5.QtGui import QIcon

APP_TITLE = "PDF 图片工具箱"
APP_VERSION = "v1.1"
GITHUB_URL = "https://github.com/xjhaz/pdf-image-toolbox"
# ========= 单位换算 =========
INCH_TO_PT = 72.0
CM_TO_PT = INCH_TO_PT / 2.54  # ≈28.3464567

def resource_path(relative_path):
    """获取资源文件路径（兼容 PyInstaller 打包后运行）"""
    if hasattr(sys, '_MEIPASS'):  # 打包后
        return os.path.join(sys._MEIPASS, relative_path)
    else:  # 源码运行
        return os.path.join(os.path.abspath("."), relative_path)

def to_pt(v: float, unit: str) -> float:
    if unit == "pt":   return float(v)
    if unit == "inch": return float(v) * INCH_TO_PT
    return float(v) * CM_TO_PT  # 默认 cm

def pt_to_unit(pt: float, unit: str) -> float:
    if unit == "pt":   return float(pt)
    if unit == "inch": return float(pt) / INCH_TO_PT
    return float(pt) / CM_TO_PT  # 默认 cm

def as_float(s: str, fb: float = 0.0) -> float:
    try: return float(str(s).replace("%", "").strip())
    except: return float(fb)

# ========= 路径统一（绝对 + 正斜杠） =========
def to_posix_abs(path: str) -> str:
    if not path: return ""
    return os.path.abspath(path).replace("\\", "/")

def resolve_posix_from_config(base_dir: str, path_in_cfg: str) -> str:
    if not path_in_cfg: return ""
    p = path_in_cfg
    if not os.path.isabs(p):
        p = os.path.join(base_dir, p)
    return to_posix_abs(p)

# ========= 通用 =========
def ensure_dir(p: str): os.makedirs(p, exist_ok=True)

def page_index(doc: fitz.Document, p: str) -> int:
    sp = (str(p) if p is not None else "last").strip().lower()
    if sp == "last": idx = len(doc) - 1
    else:
        try: idx = int(p) - 1
        except: idx = len(doc) - 1
    return max(0, min(idx, len(doc)-1))

def parse_pages(spec: str, total: int) -> List[int]:
    if not spec or not spec.strip():
        return list(range(total))
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if not part: continue
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b: a, b = b, a
            for p in range(a, b+1):
                if 1 <= p <= total: pages.add(p-1)
        else:
            try:
                p = int(part)
                if 1 <= p <= total: pages.add(p-1)
            except:
                pass
    return sorted(pages)

def rect_tuple_from_bbox(bbox) -> Optional[Tuple[float,float,float,float]]:
    if bbox is None: return None
    if hasattr(bbox, "x0") and hasattr(bbox, "y0"):
        return (float(bbox.x0), float(bbox.y0), float(bbox.x1), float(bbox.y1))
    try:
        x0, y0, x1, y1 = bbox
        return (float(x0), float(y0), float(x1), float(y1))
    except Exception:
        return None

# ========= PDF 图像与遮罩辅助 =========
def pdf_has_decode_invert(doc: fitz.Document, xref: int) -> bool:
    """读取该 XObject 的原始字典，看是否含 /Decode [1 0]（紧凑/带空格都检测）"""
    try:
        obj = doc.xref_object(xref, compressed=False)
        if obj is None: return False
        s = obj.replace(" ", "").replace("\n", "")
        # /Decode[1 0] 或 /Decode[1.0 0.0] 等都认为反相
        if "/Decode[10]" in s:  # 最常见
            return True
        # 宽松匹配，避免小数/多个空格
        if "/Decode[" in s:
            try:
                seg = s.split("/Decode[",1)[1].split("]",1)[0]
                nums = re.findall(r"[+-]?\d+(?:\.\d+)?", seg)
                if len(nums) >= 2 and float(nums[0]) == 1 and float(nums[1]) == 0:
                    return True
            except Exception:
                pass
    except Exception:
        pass
    return False

def build_pixmap_from_xref(doc: fitz.Document, xref: int) -> fitz.Pixmap:
    """
    返回一个已处理好的 Pixmap：
    - CMYK → RGB
    - 若存在 Soft Mask（/SMask），合成为 alpha，并根据 smask 的 /Decode 反相判断
    - 若主图像自身 /Decode [1 0]，对主图像反相
    """
    # 主图像
    pix = fitz.Pixmap(doc, xref)

    # CMYK → RGB
    try:
        if pix.colorspace is not None and getattr(pix.colorspace, "n", 0) == 4:
            pix = fitz.Pixmap(fitz.csRGB, pix)
    except Exception:
        pass

    # 主图像 /Decode [1 0] → 反相
    try:
        if pdf_has_decode_invert(doc, xref):
            # 仅对颜色反相；alpha（若已有）不变
            pix.invertIRect(fitz.Rect(0, 0, pix.width, pix.height))
    except Exception:
        pass

    # Soft Mask（/SMask）
    sm_xref = None
    try:
        ext = doc.extract_image(xref)  # 这一步能把 smask xref 告诉我们
        if ext and ext.get("smask"):
            sm_xref = ext["smask"]
    except Exception:
        sm_xref = None

    if sm_xref:
        try:
            sm = fitz.Pixmap(doc, sm_xref)
            # 尺寸需一致
            if sm.width != pix.width or sm.height != pix.height:
                sm = None  # 尺寸不一致则弃用
            else:
                alpha = sm.samples  # bytes
                # 若 smask 自身有 /Decode [1 0]，alpha 需要反相
                if pdf_has_decode_invert(doc, sm_xref):
                    alpha = bytes(255 - b for b in alpha)
                # 合成 alpha
                pixw = fitz.Pixmap(pix)  # 可写副本
                pixw.set_alpha(alpha)
                pix = pixw
                sm = None
        except Exception:
            pass

    return pix

# ========= 打开系统默认图片查看器 =========
def open_in_default_viewer(path: str) -> bool:
    if not path or not os.path.isfile(path): return False
    if QDesktopServices.openUrl(QUrl.fromLocalFile(path)): return True
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        return True
    except Exception:
        return False

# ========= 禁止第0列编辑（图片路径列） =========
class PathColumnNoEditDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        if index.column() == 0:
            return None
        return super().createEditor(parent, option, index)

# ========= 页签B：批量插入 =========
class TabInsert(QWidget):
    COLS = ["图片路径","X(单位)","Y(单位)","宽W(单位)","高H(单位)","X缩放%","Y缩放%","页(数字或last)","保持等比"]

    def __init__(self):
        super().__init__()
        g = QGridLayout(self); r = 0

        g.addWidget(QLabel("处理目录："), r, 0)
        self.le_root = QLineEdit()
        b_root = QPushButton("浏览…"); b_root.clicked.connect(self.choose_root)
        g.addWidget(self.le_root, r, 1, 1, 7); g.addWidget(b_root, r, 8); r += 1

        g.addWidget(QLabel("输出目录："), r, 0)
        self.le_out = QLineEdit()
        b_out = QPushButton("浏览…"); b_out.clicked.connect(self.choose_out)
        g.addWidget(self.le_out, r, 1, 1, 7); g.addWidget(b_out, r, 8); r += 1

        self.cb_suffix = QCheckBox("文件名添加后缀 _signed"); g.addWidget(self.cb_suffix, r, 1, 1, 2)

        g.addWidget(QLabel("单位："), r, 3)
        self.cb_unit = QComboBox(); self.cb_unit.addItems(["cm","pt","inch"]); self.cb_unit.setCurrentText("cm")
        g.addWidget(self.cb_unit, r, 4)

        g.addWidget(QLabel("Y坐标基准："), r, 5)
        self.cb_origin = QComboBox()
        self.cb_origin.addItems(["从下往上（PDF 标准）","从上往下（屏幕/GUI）"])
        self.cb_origin.setCurrentIndex(0)
        g.addWidget(self.cb_origin, r, 6)

        self.btn_export = QPushButton("导出配置…"); self.btn_export.clicked.connect(self.export_cfg)
        self.btn_import = QPushButton("导入配置…"); self.btn_import.clicked.connect(self.import_cfg)
        g.addWidget(self.btn_export, r, 7); g.addWidget(self.btn_import, r, 8); r += 1

        self.tab = QTableWidget(0, len(self.COLS))
        self.tab.setHorizontalHeaderLabels(self.COLS)
        self.tab.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        for c in range(1, len(self.COLS)-1):
            self.tab.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)
        self.tab.horizontalHeader().setSectionResizeMode(len(self.COLS)-1, QHeaderView.ResizeToContents)
        g.addWidget(self.tab, r, 0, 1, 9); r += 1
        self.tab.setItemDelegate(PathColumnNoEditDelegate(self.tab))

        b_add = QPushButton("添加图片…"); b_add.clicked.connect(self.add_rows)
        b_del = QPushButton("删除所选"); b_del.clicked.connect(self.del_rows)
        g.addWidget(b_add, r, 0); g.addWidget(b_del, r, 1); r += 1

        self.log = QTextEdit(); self.log.setReadOnly(True)
        g.addWidget(self.log, r, 0, 1, 9); r += 1
        self.btn_go = QPushButton("开始处理"); self.btn_go.clicked.connect(self.run)
        g.addWidget(self.btn_go, r, 0, 1, 2)

        self.tab.cellDoubleClicked.connect(self.on_cell_double_clicked)
        self.tab.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tab.customContextMenuRequested.connect(self.on_table_context_menu)

    def choose_root(self):
        d = QFileDialog.getExistingDirectory(self, "选择处理的根目录", os.getcwd())
        if d:
            self.le_root.setText(d)
            # 改动3：自动填充默认输出目录（可再次修改）
            default_out = os.path.join(d, "output")
            default_out = to_posix_abs(default_out)
            self.le_out.setText(default_out)

    def choose_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录（可不选）", os.getcwd())
        if d: self.le_out.setText(d)
    def logln(self, s: str): self.log.append(s); self.log.ensureCursorVisible()

    def add_rows(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "选择图片（可多选）", os.getcwd(),
                                                "Images (*.png *.jpg *.jpeg *.bmp *.tif *.tiff)")
        for p in paths:
            p = to_posix_abs(p)
            r = self.tab.rowCount(); self.tab.insertRow(r)
            it0 = QTableWidgetItem(p); it0.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self.tab.setItem(r, 0, it0)
            defaults = ["2.00","2.00","3.00","2.00","100","100","last"]
            for c, v in enumerate(defaults, start=1):
                it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignCenter)
                self.tab.setItem(r, c, it)
            chk = QCheckBox(); chk.setChecked(True); self.tab.setCellWidget(r, 8, chk)

    def del_rows(self):
        rows = sorted({i.row() for i in self.tab.selectedIndexes()}, reverse=True)
        for r in rows: self.tab.removeRow(r)

    def on_cell_double_clicked(self, row: int, col: int):
        if col != 0: return
        it = self.tab.item(row, 0)
        if not it: return
        path = it.text().strip()
        if not open_in_default_viewer(path):
            QMessageBox.warning(self, "无法打开", f"无法使用系统查看器打开：\n{path}")

    def on_table_context_menu(self, pos: QPoint):
        index = self.tab.indexAt(pos); row = index.row()
        if row < 0 and not self.tab.selectedIndexes(): return
        menu = QMenu(self)
        act_replace = menu.addAction("替换图片…（保持坐标不变）")
        act_delete = menu.addAction("删除所选行")
        action = menu.exec_(self.tab.viewport().mapToGlobal(pos))
        if action == act_replace: self.replace_image_for_row(row)
        elif action == act_delete: self.del_rows()

    def replace_image_for_row(self, row: int):
        if row < 0:
            rows = sorted({i.row() for i in self.tab.selectedIndexes()})
            if not rows: return
            row = rows[0]
        old = self.tab.item(row, 0).text() if self.tab.item(row, 0) else ""
        new_path, _ = QFileDialog.getOpenFileName(self, "选择替换后的图片", os.path.dirname(old) or os.getcwd(),
                                                  "Images (*.png *.jpg *.jpeg *.bmp *.tif *.tiff)")
        if not new_path: return
        new_path = to_posix_abs(new_path)
        it0 = QTableWidgetItem(new_path); it0.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
        self.tab.setItem(row, 0, it0)
        self.logln(f"🔁 第{row+1}行：已替换图片\n旧：{old}\n新：{new_path}")

    def collect_rules(self) -> List[Dict[str, Any]]:
        rules = []
        unit = self.cb_unit.currentText().strip() or "cm"
        for r in range(self.tab.rowCount()):
            def item(col):
                it = self.tab.item(r, col); return it.text().strip() if it else ""
            def chk(col):
                w = self.tab.cellWidget(r, col); return bool(w.isChecked()) if isinstance(w, QCheckBox) else False
            img = item(0)
            if not img:
                self.logln(f"⚠️ 第{r+1}行：图片路径为空，已跳过。"); continue
            X = as_float(item(1)); Y = as_float(item(2))
            W = as_float(item(3)); H = as_float(item(4))
            Sx = as_float(item(5), 100.0); Sy = as_float(item(6), 100.0)
            page = item(7) or "last"; keep = chk(8)
            if W <= 0 or H <= 0: self.logln(f"⚠️ 第{r+1}行：宽/高必须>0，已跳过。"); continue
            if Sx <= 0 or Sy <= 0: self.logln(f"⚠️ 第{r+1}行：缩放%应>0，已跳过。"); continue
            rules.append(dict(image=img, x=X, y=Y, width=W, height=H,
                              scale_x=Sx, scale_y=Sy, page=str(page), keep_aspect=keep, unit=unit))
        return rules

    def export_cfg(self):
        rules = self.collect_rules()
        if not rules:
            QMessageBox.information(self, "提示", "当前没有可导出的规则。"); return
        norm_rules = []
        for r in rules:
            r2 = dict(r); r2["image"] = to_posix_abs(r.get("image","")); norm_rules.append(r2)
        cfg = dict(
            version="1.2",
            unit=self.cb_unit.currentText().strip() or "cm",
            add_suffix=self.cb_suffix.isChecked(),
            output_dir=self.le_out.text().strip(),
            y_origin=self.cb_origin.currentText(),
            rules=norm_rules
        )
        fn, _ = QFileDialog.getSaveFileName(self, "导出配置为 JSON", os.getcwd(), "JSON (*.json)")
        if not fn: return
        try:
            with open(fn, "w", encoding="utf-8") as f: json.dump(cfg, f, ensure_ascii=False, indent=2)
            self.logln(f"✅ 已导出配置：{fn}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def import_cfg(self):
        fn, _ = QFileDialog.getOpenFileName(self, "导入配置 JSON", os.getcwd(), "JSON (*.json)")
        if not fn: return
        try:
            with open(fn, "r", encoding="utf-8") as f: cfg = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"无法读取配置：{e}"); return

        base_dir = os.path.dirname(fn)
        unit = cfg.get("unit","cm")
        self.cb_unit.setCurrentText(unit if unit in ("cm","pt","inch") else "cm")
        self.cb_suffix.setChecked(bool(cfg.get("add_suffix", False)))
        yorg = cfg.get("y_origin","从下往上（PDF 标准）")
        if yorg not in ("从下往上（PDF 标准）","从上往下（屏幕/GUI）"): yorg = "从下往上（PDF 标准）"
        self.cb_origin.setCurrentText(yorg)
        outd = cfg.get("output_dir","");  self.le_out.setText(outd or "")

        self.tab.setRowCount(0)
        for rule in cfg.get("rules", []):
            r = self.tab.rowCount(); self.tab.insertRow(r)
            it0 = QTableWidgetItem(resolve_posix_from_config(base_dir, rule.get("image","")))
            it0.setTextAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            self.tab.setItem(r, 0, it0)
            vals = [
                str(rule.get("x","")), str(rule.get("y","")),
                str(rule.get("width","")), str(rule.get("height","")),
                str(rule.get("scale_x","100")), str(rule.get("scale_y","100")),
                str(rule.get("page","last")),
            ]
            for ci, v in enumerate(vals, start=1):
                it = QTableWidgetItem(v); it.setTextAlignment(Qt.AlignCenter)
                self.tab.setItem(r, ci, it)
            chk = QCheckBox(); chk.setChecked(bool(rule.get("keep_aspect", True)))
            self.tab.setCellWidget(r, 8, chk)
        self.logln(f"✅ 已导入配置：{fn}（共{self.tab.rowCount()}条规则）")

    def run(self):
        unit = self.cb_unit.currentText().strip() or "cm"
        origin_mode = self.cb_origin.currentText()
        use_pdf_origin = origin_mode.startswith("从下往上")

        root = self.le_root.text().strip()
        if not root: QMessageBox.warning(self, "提示", "请先选择处理的根目录。"); return
        if not os.path.isdir(root): QMessageBox.warning(self, "提示", "处理目录不存在。"); return

        out_root = self.le_out.text().strip() or os.path.join(root, "output")
        out_root_abs = os.path.abspath(out_root); ensure_dir(out_root_abs)
        add_suffix = self.cb_suffix.isChecked()

        rules = self.collect_rules()
        if not rules: QMessageBox.information(self, "提示", "没有有效规则，无法处理。"); return

        self.btn_go.setEnabled(False)
        ok = 0; fail = 0
        self.logln(f"=== 开始处理（单位：{unit}；Y基准：{origin_mode}） ===")

        for dirpath, dirs, files in os.walk(root):
            abs_dirs = [os.path.abspath(os.path.join(dirpath, d)) for d in dirs]
            dirs[:] = [d for d, absd in zip(dirs, abs_dirs) if not (absd == out_root_abs or absd.startswith(out_root_abs + os.sep))]
            for f in files:
                if not f.lower().endswith(".pdf"): continue
                pdf = os.path.join(dirpath, f)

                try:
                    doc = fitz.open(pdf)
                except Exception as e:
                    fail += 1; self.logln(f"⚠️ 无法打开：{pdf} -> {e}"); continue

                if doc.is_encrypted:
                    try:
                        if not doc.authenticate(""):
                            fail += 1; self.logln(f"⚠️ 加密且无法解密，跳过：{pdf}"); doc.close(); continue
                    except Exception:
                        fail += 1; self.logln(f"⚠️ 加密文件，跳过：{pdf}"); doc.close(); continue

                try:
                    for rule in rules:
                        X = as_float(rule["x"]); Y = as_float(rule["y"])
                        W = as_float(rule["width"]); H = as_float(rule["height"])
                        Sx = as_float(rule["scale_x"], 100.0); Sy = as_float(rule["scale_y"], 100.0)
                        keep_aspect = bool(rule.get("keep_aspect", True))
                        if keep_aspect:
                            s = min(Sx, Sy) / 100.0; Wf, Hf = W * s, H * s
                        else:
                            Wf, Hf = W * (Sx/100.0), H * (Sy/100.0)

                        x_pt = to_pt(X, unit); w_pt = to_pt(Wf, unit)
                        y_input_pt = to_pt(Y, unit); h_pt = to_pt(Hf, unit)

                        pno = page_index(doc, rule["page"]); page = doc[pno]
                        page_h = page.rect.height
                        if use_pdf_origin:
                            x0 = x_pt; y0 = page_h - (y_input_pt + h_pt)
                        else:
                            x0 = x_pt; y0 = y_input_pt

                        rect = fitz.Rect(x0, y0, x0 + w_pt, y0 + h_pt)
                        page.insert_image(rect, filename=rule["image"], keep_proportion=False)

                except Exception as e:
                    fail += 1; self.logln(f"⚠️ 插入失败：{pdf} -> {e}"); doc.close(); continue

                rel = os.path.relpath(pdf, root)
                out_dir = os.path.join(out_root_abs, os.path.dirname(rel)); ensure_dir(out_dir)
                base = os.path.basename(rel)
                name, ext = (base[:-4], base[-4:]) if base.lower().endswith(".pdf") else (base, ".pdf")
                if add_suffix: name += "_signed"
                out_pdf = os.path.join(out_dir, name + ext)

                try:
                    doc.save(out_pdf); ok += 1; self.logln(f"✅ 已处理：{out_pdf}")
                except Exception as e:
                    fail += 1; self.logln(f"⚠️ 保存失败：{out_pdf} -> {e}")
                finally:
                    doc.close()

        self.logln(f"=== 完成：成功 {ok}，失败 {fail} ===")
        self.btn_go.setEnabled(True)

# ========= 页签A：从 PDF 提取 =========
class TabExtract(QWidget):
    def __init__(self):
        super().__init__()
        g = QGridLayout(self); r = 0

        g.addWidget(QLabel("PDF 文件："), r, 0)
        self.le_pdf = QLineEdit()
        b_pdf = QPushButton("浏览…"); b_pdf.clicked.connect(self.pick_pdf)
        g.addWidget(self.le_pdf, r, 1, 1, 6); g.addWidget(b_pdf, r, 7); r += 1

        g.addWidget(QLabel("导出根目录："), r, 0)
        self.le_out = QLineEdit()
        b_out = QPushButton("浏览…"); b_out.clicked.connect(self.pick_out)
        g.addWidget(self.le_out, r, 1, 1, 6); g.addWidget(b_out, r, 7); r += 1

        g.addWidget(QLabel("单位："), r, 0)
        self.cb_unit = QComboBox(); self.cb_unit.addItems(["cm","pt","inch"]); self.cb_unit.setCurrentText("cm")
        g.addWidget(self.cb_unit, r, 1)

        g.addWidget(QLabel("Y坐标基准："), r, 2)
        self.cb_origin = QComboBox()
        self.cb_origin.addItems(["从下往上（PDF 标准）","从上往下（屏幕/GUI）"])
        self.cb_origin.setCurrentIndex(0)
        g.addWidget(self.cb_origin, r, 3)

        g.addWidget(QLabel("页码（如1,3-5）："), r, 4)
        self.le_pages = QLineEdit()
        g.addWidget(self.le_pages, r, 5, 1, 2); r += 1

        self.cb_flatten = QCheckBox("导出时白底（去透明）")
        self.cb_flatten.setChecked(False)
        g.addWidget(self.cb_flatten, r, 1, 1, 3); r += 1

        self.log = QTextEdit(); self.log.setReadOnly(True)
        g.addWidget(self.log, r, 0, 1, 8); r += 1

        self.btn_go = QPushButton("扫描并导出")
        self.btn_go.clicked.connect(self.scan_and_export)
        g.addWidget(self.btn_go, r, 0, 1, 2)

    def logln(self, s: str): self.log.append(s); self.log.ensureCursorVisible()

    def pick_pdf(self):
        fn, _ = QFileDialog.getOpenFileName(self, "选择 PDF 文件", os.getcwd(), "PDF (*.pdf)")
        if fn:
            self.le_pdf.setText(fn)
            # 改动2：自动将导出根目录填成 “<pdf同目录>\pic”（用户可再次修改）
            pdf_dir = os.path.dirname(fn)
            default_out = os.path.join(pdf_dir, "pic")
            default_out = to_posix_abs(default_out)
            self.le_out.setText(default_out)

    def pick_out(self):
        d = QFileDialog.getExistingDirectory(self, "选择导出根目录", os.getcwd())
        if d: self.le_out.setText(d)

    def scan_and_export(self):
        pdf_path = self.le_pdf.text().strip()
        if not pdf_path or not os.path.isfile(pdf_path):
            QMessageBox.warning(self, "提示", "请先选择一个有效的 PDF 文件。"); return

        unit = self.cb_unit.currentText().strip() or "cm"
        origin_mode = self.cb_origin.currentText()
        use_pdf_origin = origin_mode.startswith("从下往上")

        # 改动2：导出根目录默认就是 <pdf同目录>\pic，并直接在该目录下保存图片与 JSON
        out_root = self.le_out.text().strip()
        if not out_root:
            out_root = os.path.join(os.path.dirname(pdf_path), "pic")
        out_root = os.path.abspath(out_root)
        ensure_dir(out_root)

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开PDF：{e}"); return

        try:
            if doc.is_encrypted:
                try:
                    if not doc.authenticate(""): QMessageBox.critical(self, "错误", "PDF 已加密且无法解密。"); return
                except Exception:
                    QMessageBox.critical(self, "错误", "PDF 已加密且无法解密。"); return

            total = len(doc)
            pages = parse_pages(self.le_pages.text(), total) or list(range(total))

            self.btn_go.setEnabled(False)
            self.logln(f"=== 开始扫描 ===")
            self.logln(f"PDF：{pdf_path}")
            self.logln(f"导出根目录：{out_root}")
            self.logln(f"单位：{unit}；Y基准：{origin_mode}；页码：{', '.join(str(p+1) for p in pages)}")

            rules: List[Dict[str, Any]] = []
            img_count = 0
            pdf_base = os.path.splitext(os.path.basename(pdf_path))[0]

            for pno in pages:
                page = doc[pno]; page_h = page.rect.height

                # 找出本页所有图片的 xref 与矩形
                items: List[Tuple[int, Tuple[float,float,float,float]]] = []
                used_new = False
                if hasattr(page, "get_image_info"):
                    try:
                        infos = page.get_image_info(xrefs=True); used_new = True
                        for info in infos:
                            rt = rect_tuple_from_bbox(info.get("bbox"))
                            xref = info.get("xref") or info.get("image") or info.get("xref0")
                            if rt and xref: items.append((int(xref), rt))
                    except Exception:
                        used_new = False
                if not used_new:
                    for img in page.get_images(full=True):
                        xref = img[0]
                        rects = []
                        try: rects = page.get_image_rects(xref)
                        except Exception: rects = []
                        for rr in rects:
                            rt = rect_tuple_from_bbox(rr)
                            if rt: items.append((xref, rt))

                # 逐个原图导出（合成 alpha / 反相修正）
                for xref, rect in items:
                    x0, y0, x1, y1 = rect
                    w_pt, h_pt = x1 - x0, y1 - y0

                    try:
                        pix = build_pixmap_from_xref(doc, xref)
                        if self.cb_flatten.isChecked() and pix.alpha:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                    except Exception as e:
                        self.logln(f"⚠️ 第{pno+1}页 xref={xref} 提取失败 -> {e}")
                        continue

                    img_count += 1
                    img_name = f"{pdf_base}_{img_count:04d}.png"
                    img_path = os.path.join(out_root, img_name)  # 改动2：直接导出到 out_root
                    try:
                        pix.save(img_path)
                    except Exception as e:
                        self.logln(f"⚠️ 第{pno+1}页 保存 PNG 失败：{img_name} -> {e}")
                        continue
                    finally:
                        try: pix = None
                        except: pass

                    # 导出给插入器的坐标（单位换算；Y基准=“从上往下量”）
                    if use_pdf_origin:
                        X_unit = pt_to_unit(x0, unit)
                        Y_unit = pt_to_unit(page_h - y1, unit)  # 左下原点 → 从上往下量
                    else:
                        X_unit = pt_to_unit(x0, unit)
                        Y_unit = pt_to_unit(y0, unit)

                    W_unit = pt_to_unit(w_pt, unit)
                    H_unit = pt_to_unit(h_pt, unit)

                    rule = dict(
                        image=to_posix_abs(img_path),
                        x=round(X_unit, 4),
                        y=round(Y_unit, 4),
                        width=round(W_unit, 4),
                        height=round(H_unit, 4),
                        scale_x=100.0,
                        scale_y=100.0,
                        page=str(pno + 1),
                        keep_aspect=True,
                        unit=unit,
                    )
                    rules.append(rule)

                    self.logln(
                        f"第{pno+1}页：保存 {img_name} | "
                        f"X={rule['x']}{unit}, Y={rule['y']}{unit}, "
                        f"W={rule['width']}{unit}, H={rule['height']}{unit}"
                    )

            cfg = dict(
                version="1.2",
                unit=unit,
                add_suffix=False,   # 提取配置默认不加后缀
                output_dir="",      # 导出的插入配置不强制指定输出目录
                y_origin="从下往上（PDF 标准）" if use_pdf_origin else "从上往下（屏幕/GUI）",
                rules=rules,
            )
            # 改动1：JSON 配置名改为 “PDF名 + _config.json”
            json_path = os.path.join(out_root, f"{pdf_base}_config.json")
            try:
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(cfg, f, ensure_ascii=False, indent=2)
                self.logln(f"=== 完成：导出图片 {img_count} 个 ===")
                self.logln(f"JSON 配置：{json_path}")
                QMessageBox.information(self, "完成",
                    f"已导出 {img_count} 张 PNG 到\n{out_root}\n并生成配置：\n{json_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"写入 JSON 失败：{e}")
        finally:
            try: doc.close()
            except Exception: pass
            self.btn_go.setEnabled(True)

class TabAbout(QWidget):
    def __init__(self):
        super().__init__()
        g = QGridLayout(self); r = 0

        title = QLabel(
            f"<h2 style='margin:0;'>PDF Image Toolbox "
            f"<span style='font-size:12px; color:#666; border:1px solid #ddd; "
            f"padding:2px 8px; border-radius:10px; vertical-align:middle;'>{APP_VERSION}</span>"
            f"</h2>"
        )
        g.addWidget(title, r, 0, 1, 2); r += 1

        desc = QLabel("基于 <b>PyQt5</b> 与 <b>PyMuPDF (fitz)</b> 的 PDF 图像提取 / 批量插入工具。")
        desc.setWordWrap(True)
        g.addWidget(desc, r, 0, 1, 2); r += 1

        link = QLabel(f"GitHub：<a href='{GITHUB_URL}'>{GITHUB_URL}</a>")
        link.setOpenExternalLinks(True)
        g.addWidget(link, r, 0, 1, 2); r += 1

        btn_open = QPushButton("打开 GitHub")
        btn_copy = QPushButton("复制仓库地址")
        btn_open.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(GITHUB_URL)))
        btn_copy.clicked.connect(lambda: self._copy(GITHUB_URL))
        g.addWidget(btn_open, r, 0)
        g.addWidget(btn_copy, r, 1); r += 1

        note = QLabel("© 2025 xjhaz")
        note.setStyleSheet("color:#888;")
        g.addWidget(note, r, 0, 1, 2)

    def _copy(self, text: str):
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "已复制", "仓库地址已复制到剪贴板。")

# ========= 主窗口 =========
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE); self.resize(1220, 780)
        self.setWindowIcon(QIcon(resource_path("pdf_toolbox.ico")))
        tabs = QTabWidget()
        self.tab_insert = TabInsert()
        self.tab_extract = TabExtract()
        self.tab_about  = TabAbout() 
        tabs.addTab(self.tab_extract, "从PDF提取配置")
        tabs.addTab(self.tab_insert, "批量插入")
        tabs.addTab(self.tab_about,  "关于")  
        self.setCentralWidget(tabs)
        

def main():
    app = QApplication(sys.argv)
    w = MainWindow(); w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
