"""Headless recalculation with LibreOffice Calc over UNO.

Approach (tested on LibreOffice 24.2, see README): start `soffice --headless` on a private profile and port,
load the workbook through the Python `uno` bridge, call `calculateAll()` and store the result with the
"Calc MS Excel 2007 XML" filter. Formulas are kept and the cached values are LibreOffice's, so
`openpyxl.load_workbook(out, data_only=True)` reads the recalculated numbers (XIRR / XNPV included).

Fallback when `uno` is not importable: `soffice --headless --convert-to xlsx` also recalculates on load
(`--infilter` is not needed for xlsx) but cannot run macros or set inputs.
"""
from __future__ import annotations

import os
import shutil
import socket
import subprocess
import tempfile
import time
from pathlib import Path

XLSX_FILTER = "Calc MS Excel 2007 XML"
XLSM_FILTER = "Calc MS Excel 2007 VBA XML"


def soffice() -> str:
    exe = os.environ.get("SOFFICE") or shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise RuntimeError("LibreOffice 'soffice' not found (set SOFFICE)")
    return exe


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _prop(name, value):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


class Office:
    """One soffice process with one loaded workbook (a private copy). Use as a context manager."""

    def __init__(self, path: str | os.PathLike):
        self.src = Path(path).resolve()
        self.tmp = tempfile.mkdtemp(prefix="fepm_lo_")
        self.proc = self.desktop = self.doc = None

    def __enter__(self):
        import uno
        port = _free_port()
        accept = f"socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        profile = uno.systemPathToFileUrl(os.path.join(self.tmp, "profile"))
        env = dict(os.environ, HOME=self.tmp)
        self.proc = subprocess.Popen(
            [soffice(), "--headless", "--invisible", "--norestore", "--nologo", "--nodefault", "--nolockcheck",
             f"-env:UserInstallation={profile}", f"--accept={accept}"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, env=env, start_new_session=True)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        t0, ctx = time.time(), None
        while ctx is None:
            try:
                ctx = resolver.resolve(f"uno:{accept}")
            except Exception:
                if time.time() - t0 > 60 or self.proc.poll() is not None:
                    self.__exit__()
                    raise RuntimeError("could not connect to soffice")
                time.sleep(0.25)
        self.ctx = ctx
        self.desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        work = os.path.join(self.tmp, self.src.name)
        shutil.copyfile(self.src, work)
        # MacroExecutionMode 4 = always execute, no warning (needed for run_macro; nothing auto-runs on load).
        self.doc = self.desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(work), "_blank", 0, (_prop("Hidden", True), _prop("MacroExecutionMode", 4)))
        if self.doc is None:
            self.__exit__()
            raise RuntimeError(f"LibreOffice could not load {self.src}")
        self.calculate()
        return self

    def __exit__(self, *exc):
        for step in (lambda: self.doc.close(True), lambda: self.desktop.terminate()):
            try:
                step()
            except Exception:
                pass
        if self.proc:
            try:
                self.proc.wait(timeout=20)
            except subprocess.TimeoutExpired:
                self.proc.kill()
        shutil.rmtree(self.tmp, ignore_errors=True)

    # ---- cells -------------------------------------------------------------------------------
    def calculate(self):
        self.doc.calculateAll()

    def _range(self, ref: str):
        """`Sheet!A1`, `Sheet!A1:B2` or a workbook name."""
        if "!" in ref:
            sheet, addr = ref.split("!")
            return self.doc.Sheets.getByName(sheet.strip("'")).getCellRangeByName(addr.replace("$", ""))
        return self.doc.NamedRanges.getByName(ref).getReferredCells()

    def get(self, ref: str):
        cell = self._range(ref).getCellByPosition(0, 0)
        if cell.getError():
            return f"#ERR{cell.getError()}"
        t = cell.getType().value
        if t == "EMPTY":
            return None
        if t == "TEXT" or (t == "FORMULA" and cell.FormulaResultType2 == 2):
            return cell.getString()
        return cell.getValue()

    def set(self, ref: str, value):
        cell = self._range(ref).getCellByPosition(0, 0)
        cell.setString(value) if isinstance(value, str) else cell.setValue(value)

    # ---- macros ------------------------------------------------------------------------------
    def add_vba_modules(self, modules: dict[str, str]):
        """Insert VBA source as standard modules of the document's Standard library (VBA compat mode)."""
        import uno
        from com.sun.star.script import ModuleInfo
        libs = self.doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        libs.VBACompatibilityMode = True
        libs.ProjectName = "Standard"  # the VBA exporter looks the library up by project name
        lib = libs.getByName("Standard")
        info = ModuleInfo()
        info.ModuleType = uno.getConstantByName("com.sun.star.script.ModuleType.NORMAL")
        for name, src in modules.items():
            # LibreOffice writes its own "Attribute VB_Name" line on export; drop the source's attributes.
            src = "\n".join(ln for ln in src.splitlines() if not ln.startswith("Attribute "))
            src = src if "Option VBASupport" in src else "Option VBASupport 1\n" + src
            if lib.hasByName(name):
                lib.removeByName(name)
            if not lib.hasModuleInfo(name):
                lib.insertModuleInfo(name, info)  # without ModuleInfo the xlsm export fails
            lib.insertByName(name, src)

    def _config(self, node: str, prop: str, value):
        cp = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", self.ctx)
        acc = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess",
                                             (_prop("nodepath", node),))
        acc.setPropertyValue(prop, value)
        acc.commitChanges()

    def run_macro(self, module: str, proc: str):
        uri = f"vnd.sun.star.script:Standard.{module}.{proc}?language=Basic&location=document"
        return self.doc.getScriptProvider().getScript(uri).invoke((), (), ())

    def save(self, out: str | os.PathLike, macro: bool = False):
        import uno
        out = Path(out).resolve()
        if macro:  # LibreOffice's (experimental) VBA exporter writes the Basic modules into vbaProject.bin
            self._config("/org.openoffice.Office.Calc/Filter/Import/VBA", "UseExport", True)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.doc.storeToURL(uno.systemPathToFileUrl(str(out)),
                            (_prop("FilterName", XLSM_FILTER if macro else XLSX_FILTER),))
        return out


def uno_available() -> bool:
    try:
        import uno  # noqa: F401
        return True
    except ImportError:
        return False


def recalc(path: str | os.PathLike, out: str | os.PathLike | None = None,
           inputs: dict | None = None) -> Path:
    """Recalculate `path` and write the result (formulas + LibreOffice values) to `out`.

    `inputs` maps workbook names or `Sheet!A1` to values to set before recalculating."""
    path = Path(path)
    out = Path(out) if out else path.with_name(path.stem + "_recalc" + path.suffix)
    if uno_available():
        with Office(path) as lo:
            for ref, v in (inputs or {}).items():
                lo.set(ref, v)
            lo.calculate()
            lo.save(out, macro=out.suffix == ".xlsm")
        return out
    if inputs:
        raise RuntimeError("setting inputs needs the Python uno bridge")
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([soffice(), "--headless", f"-env:UserInstallation=file://{td}/p", "--convert-to", "xlsx",
                        "--outdir", td, str(path)], check=True, capture_output=True, timeout=300)
        shutil.move(os.path.join(td, path.stem + ".xlsx"), out)
    return out


def values(path: str | os.PathLike):
    """openpyxl workbook with cached (recalculated) values."""
    import openpyxl
    return openpyxl.load_workbook(path, data_only=True)


def name_value(wb_values, name: str):
    """Value of a single-cell workbook name in a data_only workbook."""
    dn = wb_values.defined_names[name]
    sheet, cell = next(dn.destinations)
    return wb_values[sheet][cell.replace("$", "")].value
