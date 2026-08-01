"""converter.py —— docx → PDF 多引擎转换（真正按序降级）。

引擎优先级：Microsoft Word COM → WPS (KWPS) COM → LibreOffice headless。
- Word/WPS 用 DispatchEx 创建独立实例，绝不触碰用户正在使用的 Office
- 引擎失败自动尝试下一个；全部失败返回各引擎错误摘要
- 全局锁串行化转换（Office 单实例约束）
"""

from __future__ import annotations

import functools
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time

# Word/WPS 导出 PDF 的 FileFormat 常量（wdFormatPDF = 17）
_WD_FORMAT_PDF = 17
_LOCK = threading.Lock()

_LIBREOFFICE_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def _find_soffice() -> str | None:
    for p in _LIBREOFFICE_PATHS:
        if os.path.exists(p):
            return p
    return shutil.which("soffice") or shutil.which("libreoffice")


@functools.lru_cache(maxsize=1)
def detect_engines() -> list[str]:
    """探测可用引擎，按优先级返回列表（word/wps/libreoffice）。"""
    engines = []
    if sys.platform == "win32":
        if _com_engine_available("Word.Application"):
            engines.append("word")
        if _com_engine_available("KWPS.Application"):
            engines.append("wps")
    if _find_soffice():
        engines.append("libreoffice")
    return engines


def _com_engine_available(prog_id: str) -> bool:
    """用 DispatchEx 起临时实例探测引擎（不触碰用户已有实例），用完即退。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    hr = pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx(prog_id)
        app.Visible = False
        return True
    except Exception:  # noqa: BLE001
        return False
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass
        if hr == 0:
            pythoncom.CoUninitialize()


def _convert_via_com(docx_path: str, pdf_path: str, prog_id: str) -> None:
    """通过 Word/WPS COM 转换（DispatchEx 独立实例，防误关用户文档）。"""
    import pythoncom
    import win32com.client

    hr = pythoncom.CoInitialize()
    app = None
    doc = None
    try:
        app = win32com.client.DispatchEx(prog_id)
        app.Visible = False
        app.DisplayAlerts = 0
        doc = app.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=_WD_FORMAT_PDF)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:  # noqa: BLE001
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass
        if hr == 0:
            pythoncom.CoUninitialize()


def _kill_process_tree(proc) -> None:
    """终止 soffice 进程树（Windows 用 taskkill /T，其他平台 killpg 兜底）。"""
    try:
        if sys.platform == "win32":
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                           capture_output=True, timeout=15)
        else:
            proc.kill()
    except Exception:  # noqa: BLE001
        try:
            proc.kill()
        except Exception:  # noqa: BLE001
            pass


def _convert_via_libreoffice(docx_path: str, pdf_path: str) -> None:
    """通过 LibreOffice headless 转换（独立 profile 防并发锁冲突）。"""
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice")
    outdir = os.path.dirname(os.path.abspath(pdf_path))
    os.makedirs(outdir, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    cmd = [soffice, "--headless", "-env:UserInstallation=file:///" + profile.replace("\\", "/"),
           "--convert-to", "pdf", "--outdir", outdir, os.path.abspath(docx_path)]
    kwargs = {"capture_output": True, "timeout": 120}
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    proc = subprocess.run(cmd, **kwargs)
    if proc.returncode != 0:
        raise RuntimeError(f"LibreOffice 转换失败: {proc.stderr.decode(errors='replace')[:200]}")
    generated = os.path.join(outdir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if not os.path.exists(generated):
        raise RuntimeError("LibreOffice 未生成 PDF 文件")
    if os.path.abspath(generated) != os.path.abspath(pdf_path):
        os.replace(generated, pdf_path)


def convert(docx_path: str, output_path: str | None = None) -> dict:
    """docx → PDF。按 word → wps → libreoffice 顺序尝试，失败自动降级。"""
    if not os.path.exists(docx_path):
        return {"ok": False, "error": f"源文件不存在: {docx_path}"}
    if not docx_path.lower().endswith(".docx"):
        return {"ok": False, "error": "仅支持 .docx 转换"}

    if output_path is None:
        output_path = os.path.splitext(docx_path)[0] + ".pdf"
    if not output_path.lower().endswith(".pdf"):
        return {"ok": False, "error": "输出路径必须以 .pdf 结尾"}
    outdir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(outdir, exist_ok=True)
    # 转换前删除陈旧目标文件，避免失败时误报成功
    if os.path.exists(output_path):
        os.remove(output_path)

    engines = detect_engines()
    if not engines:
        return {"ok": False, "error": "未检测到可用 PDF 引擎（需安装 Word/WPS/LibreOffice 之一）"}

    errors = []
    with _LOCK:  # Office 单实例，全局串行化
        for engine in engines:
            start = time.time()
            try:
                if engine == "word":
                    _convert_via_com(docx_path, output_path, "Word.Application")
                elif engine == "wps":
                    _convert_via_com(docx_path, output_path, "KWPS.Application")
                else:
                    _convert_via_libreoffice(docx_path, output_path)
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    return {"ok": True, "path": output_path, "engine": engine,
                            "elapsed_seconds": round(time.time() - start, 2)}
                errors.append(f"{engine}: 未生成有效 PDF")
            except subprocess.TimeoutExpired:
                errors.append("libreoffice: 转换超时(120s)")
            except Exception as e:  # noqa: BLE001
                errors.append(f"{engine}: {type(e).__name__}: {str(e)[:120]}")
    return {"ok": False, "error": "所有引擎均失败", "engine_errors": errors}
