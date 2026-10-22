# v4.5（5a-1）：以 LibreOffice（UNO）驅動 Excel 活頁簿——建置時的 Python 求解與快照共用（反向 DCF；5a 另有敏感度快照）。
# 用法：
#   with Workbook('檔案.xlsx') as wb:
#       c = wb.cell('輸入與假設', '穩態 EBITDA 率（FY30）')   # 依 A 欄列名稱找儲存格（預設 C 欄）
#       wb.set(c, 0.6); wb.get(wb.cell('評價_DCF與目標價', '加權目標價'))
# 改輸入後 LibreOffice 自動重算，讀值即為重算後的結果；with 區塊結束時一律關閉文件、不存檔（除非呼叫 save()）。
import os, subprocess, tempfile, time, socket, shutil
from pathlib import Path


def _free_port():
    s = socket.socket(); s.bind(('127.0.0.1', 0)); p = s.getsockname()[1]; s.close(); return p


class Workbook:
    def __init__(self, path, timeout=60):
        self.path, self.timeout = os.path.abspath(path), timeout

    def __enter__(self):
        import uno
        from com.sun.star.beans import PropertyValue
        self.prof = tempfile.mkdtemp(prefix='lo_uno_')
        port = _free_port()
        env = dict(os.environ, LC_ALL='C.UTF-8', LANG='C.UTF-8')  # 與 recalc.py 相同：避免中文檔名變成 ??
        self.proc = subprocess.Popen(['soffice', '-env:UserInstallation=' + Path(self.prof).as_uri(), '--headless', '--norestore',
                                      '--nologo', f'--accept=socket,host=127.0.0.1,port={port};urp;'],
                                     env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
        ctx = uno.getComponentContext()
        res = ctx.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', ctx)
        t0 = time.time()
        while True:
            try:
                rc = res.resolve(f'uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext'); break
            except Exception:
                if time.time() - t0 > self.timeout: self._kill(); raise RuntimeError('LibreOffice 無法啟動（UNO 連線逾時）')
                time.sleep(0.3)
        desk = rc.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop', rc)
        pv = PropertyValue(); pv.Name = 'Hidden'; pv.Value = True
        self.doc = desk.loadComponentFromURL(Path(self.path).as_uri(), '_blank', 0, (pv,))
        if self.doc is None: self._kill(); raise RuntimeError(f'LibreOffice 開不了檔案：{self.path}')
        self.doc.calculateAll()
        return self

    def cell(self, sheet, label, col='C', need_value=True, prefix=False):
        """A 欄等於 label（prefix＝True 時為以 label 開頭）的第一列（need_value：該列 col 欄有數值或公式），回傳 (工作表, 'C12')。找不到即報錯，不猜。"""
        sh = self.doc.Sheets.getByName(sheet)
        cur = sh.createCursor(); cur.gotoEndOfUsedArea(False); last = cur.RangeAddress.EndRow + 1
        for r in range(1, last + 1):
            a = sh.getCellRangeByName(f'A{r}').getString()
            if a == label or (prefix and a.startswith(label)):
                c = sh.getCellRangeByName(f'{col}{r}')
                if not need_value or c.getFormula() != '': return (sh, f'{col}{r}')
        raise KeyError(f'找不到：{sheet}｜{label}（{col} 欄）')

    def get(self, ref):
        sh, a = ref; c = sh.getCellRangeByName(a)
        if c.getError(): raise ValueError(f'{a} 公式錯誤 {c.getError()}')
        return c.getValue()

    def set(self, ref, x):
        sh, a = ref; sh.getCellRangeByName(a).setValue(x)

    def save(self):
        self.doc.store()

    def _kill(self):
        try: os.killpg(self.proc.pid, 9)
        except Exception: pass
        shutil.rmtree(self.prof, ignore_errors=True)

    def __exit__(self, *a):
        try: self.doc.close(True)
        except Exception: pass
        self._kill()
