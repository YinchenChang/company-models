# 外部稽核優化 W3-4（2026-11-23）：NVIDIA「執行長異動」只進 Sens 與風險登錄，基準不動（Andy 2026-11-20 決定 J1，取預設 (a)）。
# 依據：docs/workorders/20261120_外部稽核優化_工作單.md W3-4；docs/plan/高不確定性預估工作原則.md（Assumed 不當點值入基準，只放 Sens）。
# 做什麼：
#   1. Sens 頁末追加「執行長異動敏感度」區塊：三案同時成立（整本重算快照，規則 X6）、各案評等、檢驗指標；單案數字不重複列（見上方業務主因表，#716）。數字由 nvidia/sens_nvidia.py 整本重建、
#      LibreOffice 重算後寫入 nvidia/sens.json 的 "ceo"；本檔只把快照寫進 Sens 頁。
#   2. README 頁風險登錄的「執行長異動」列（列位由 v035_post.py 的佔位文字決定）改寫成帶 Sens 數字的公式。
# 性質：只新增（Sens 頁末；README 風險登錄多一列，其後各列下移一列）。既有數值格、基準目標價、評等不變；不升版、不加沿革列（主控統一處理）。
# 三個案例的參數值全部是 Assumed：來自 Gemini 試算（2026-11-19），沒有實證，只作敏感度，基準不採。本檔不自創其他參數。
import json, os
from openpyxl.styles import Alignment, Font, PatternFill

BOLD = Font(bold=True); HDR = PatternFill("solid", fgColor="DDEBF7"); WRAP = Alignment(wrap_text=True, vertical="top")
SRC = "來自 Gemini 試算（2026-11-19）與提案 G-T05（2026-11-23），無實證，只作敏感度；基準不採"
WT = "tp_wA=0.15;tp_wB=0.55;tp_wC=0.3"

# （鍵, 案例名, 設定, NV_OVERRIDE；None＝基準列，不重建）。sens_nvidia.py 也讀這張表，案例只在這裡定義一次。
# 整合（2026-11-24）：單案的加權目標價只列在上方「業務主因」表（v035_post.ceo_cases 的活頁簿公式列；#716，Gemini G-T05 的刻度），
# 本區塊不重複列單案數字，只留那張表沒有的內容：三案同時成立的整本重算、各案評等、檢驗指標。單案仍整本重建，用途是取評等與事後交叉核對公式列。
CEO_CASES = [
    ("base", "基準（現行模型）", "相對倍數溢價保留 100%、WACC +0bp、情境權重 A／B／C 30／50／20", None),
    ("rm50", "相對倍數溢價保留 50%", "Comparable 的相對倍數調整係數只保留超出 1 的一半：係數＝1＋0.5 ×（原係數 − 1）；DCF 不變", "_rmkeep=0.5"),
    ("w100", "WACC 加 100bp", "In 頁 wacc_adj＝+1%（RONIC 隨 WACC）；Comparable 不變", "wacc_adj=0.01"),
    ("wt", "情境權重改 15／55／30", "In 頁 tp_wA／tp_wB／tp_wC＝0.15／0.55／0.30（A 樂觀、B 基準、C 保守；現行 0.30／0.50／0.20）；DCF 與 Comparable 的前瞻 EPS 都用新權重", WT),
    ("all", "三案同時成立", "相對倍數溢價保留 50%＋WACC 加 100bp＋情境權重 15／55／30", "_rmkeep=0.5;wacc_adj=0.01;" + WT),
]
KEYCOL = {"base": "", "rm50": "rm_factor", "w100": "wacc_adj", "wt": "tp_wA、tp_wB、tp_wC", "all": "rm_factor、wacc_adj、tp_wA、tp_wB、tp_wC"}
# 單案數字所在：（業務主因表／目標價參數表的列標籤開頭, 欄, 給讀者看的位置說明）
WHERE = {
    "rm50": ("執行長異動：相對倍數溢價保留 50%", "C", "上方業務主因表「執行長異動：相對倍數溢價保留 50%」列的低案"),
    "w100": ("WACC ±1pt", "E", "上方目標價參數表「WACC ±1pt」列的高案"),
    "wt": ("執行長交接期情境權重偏保守 15／55／30", "C", "上方業務主因表「執行長交接期情境權重偏保守 15／55／30」列的低案"),
}
NOTE = {
    "rm50": SRC + "。50% 這個比例沒有歷史樣本或同業案例支持；25%、75% 與 0% 的刻度見業務主因表",
    "w100": SRC + "。本模型 WACC 沒有公司特有風險項，100bp 是直接加在 WACC 上的假設；＋0.5pt、＋2pt 的刻度見業務主因表",
    "wt": SRC + "。15／55／30 依樂觀／基準／保守（A／B／C）順序，Gemini G-T05 已確認；10／45／45、20／55／25 的刻度見業務主因表",
    "all": SRC + "。三項互相不獨立（都在表達同一個擔憂），合併會重複計算，只供看量級",
}
# 「未來兩季可觀察、能檢驗它的指標」：先寫模型內已有、有資料來源的指標；模型內沒有的寫「未找到」，另列 Gemini 提案的外部指標並標明未驗證，不硬湊。
IND_RM = ("Val 頁 0.6 節「現行相對倍數」（現價 ÷ 12 個月前瞻 EPS ÷ 同業前瞻 P/E 中位數）在 FY27 Q4（約 2027-02）、FY28 Q1（約 2027-05）財報後的讀數，"
          "對照 10 年比值（調整係數）。只能檢驗市場有沒有給溢價，分不出原因是不是執行長異動；把異動對應到「溢價剩一半」的實證：未找到。"
          "Gemini G-T05 提案的門檻（財報後 20 日均值低於 1.20 倍或高於 1.55 倍）是 Gemini 的假設，未驗證")
IND_W = ("模型內：未找到。本模型 WACC 的 β 是同業中位數的自算產業 β，rf、ERP 是市場層輸入，沒有公司特有風險項。"
         "Gemini G-T05 提案三個外部指標（10 年期公司債利差較異動前走闊 25bp 以上；8-K Item 5.02 正式任命時程與具名主管離職達 2 人；"
         "3 個月隱含波動率相對費城半導體指數高出 10pt 以上）：門檻是 Gemini 的假設，本模型沒有收錄這些資料來源，未驗證")
IND_WT = ("FY27 Q4 財報（約 2027-02）：營收對公司指引中點（Hist 頁 h_q4g_rev）、年底存貨落在全額執行或趨勢（Recon I2 預登錄檢驗點）；"
          "FY28 Q1 財報（約 2027-05）：公司是否維持 FY2028 營收成長展望（Hist 頁 h_g28_co）。只能檢驗各情境的相對支持度；把異動對應到 15／55／30 這組權重的實證：未找到")
IND = {"rm50": IND_RM, "w100": IND_W, "wt": IND_WT, "all": "同上三案各自的指標；合併本身沒有另外的檢驗指標（未找到）"}
HEAD = ["案例", "設定", "加權目標價", "相對基準", "DCF 目標價", "Comparable 目標價（相對倍數調整後）", "基準加權目標價", "鍵", "上檔空間", "評等", "標記", "說明與來源", "未來兩季可觀察、能檢驗它的指標"]
TITLE = "執行長異動敏感度"


def _find(ws, prefix):
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.lstrip("=\"").startswith(prefix):
            return r
    raise KeyError(prefix)


def apply(wb, here):
    SE, RD = wb["Sens"], wb["README"]
    p = os.path.join(here, "sens.json")
    SJ = json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}
    C = SJ.get("ceo") or {}
    b = C.get("base") or {}
    ok = bool(b) and abs(b.get("tp", 0) - (SJ.get("_base") or {}).get("tp", -1)) < 5e-10   # 快照基準須與 sens.json 的基準檔同一次重建
    src = {k: (_find(SE, lab), col) for k, (lab, col, _) in WHERE.items()}   # 單案數字的唯一位置（上方表的公式列）

    r = SE.max_row + 3   # 與上方 HTML 主圖用表隔兩列空白（該表讀到空白列為止）
    t0 = r
    SE.cell(r, 1, f"{TITLE}：三案同時成立、各案評等與檢驗指標（外部稽核優化 W3-4，2026-11-23；Andy 2026-11-20 決定：執行長異動只進 Sens 與風險登錄，基準不動）").font = BOLD; r += 1
    SE.cell(r, 1, "三個案例的參數值全部標 Assumed：" + SRC + "。單案的加權目標價只列在上方「業務主因」表與目標價參數表（活頁簿公式），本區塊不重複列；"
                  "本區塊補那兩張表沒有的三項：三案同時成立的整本重算、各案評等、檢驗指標。「三案同時成立」一列與各案評等是整本重算快照"
                  "（nvidia/sens_nvidia.py → nvidia/sens.json 的 ceo；規則 X6），不是活頁簿公式；In 頁改過須重跑（模型狀態見本頁「模型治理」）。"
                  "事件本身的一手來源本模型沒有收錄。"
                  "對照（主控轉述，未見原檔）：Gemini 以舊版舊價試算三案合併為 329.77 → 249.69；本表以現行版重算，兩者只比量級。"); r += 1
    for j, h in enumerate(HEAD, 1):
        c = SE.cell(r, j, h); c.font = BOLD; c.fill = HDR
    r += 1
    rows = {}
    for k, name, setting, ov in CEO_CASES:
        if not ov: continue   # 基準列不列（基準加權目標價見「三案同時成立」列的對照欄）
        d = C.get(k) if ok else None
        SE.cell(r, 1, name); SE.cell(r, 2, setting); SE.cell(r, 8, KEYCOL[k]); SE.cell(r, 11, "Assumed"); SE.cell(r, 12, NOTE[k]); SE.cell(r, 13, IND[k])
        if k in WHERE:
            SE.cell(r, 3, "見" + WHERE[k][2])
            if d: SE.cell(r, 10, d["rec"])
        elif d:
            SE.cell(r, 3, d["tp"]).number_format = "#,##0.00"; SE.cell(r, 4, d["tp"] / b["tp"] - 1).number_format = "+0.0%;-0.0%"
            SE.cell(r, 5, d["dcf"]).number_format = "#,##0.00"; SE.cell(r, 6, d["comp"]).number_format = "#,##0.00"
            SE.cell(r, 7, b["tp"]).number_format = "#,##0.00"; SE.cell(r, 9, d["up"]).number_format = "+0.0%;-0.0%"; SE.cell(r, 10, d["rec"])
        else:
            SE.cell(r, 3, "—")
        rows[k] = r; r += 1
    SE.cell(r, 1, "讀法：「三案同時成立」是把三項疊在一起整本重算，不是三個單案變動的加總。"
                  "三項都沒有實證，也彼此不獨立，所以不進基準、不改情境；有新證據時（上表最右欄的指標）再重判。")

    # README 風險登錄：把佔位列改寫成帶數字的公式（列號不變）。單案數字指向上方表的公式列，三案合併指向本區塊。
    rr = _find(RD, "執行長異動")
    if ok:
        S = lambda k: f'TEXT(Sens!{src[k][1]}{src[k][0]},"0.00")'
        A = rows["all"]
        RD.cell(rr, 1).value = (
            '="執行長異動：未建模於基準（Andy 2026-11-20 決定：只進 Sens 與風險登錄，基準目標價、情境與評等不因此改變）。事件本身的一手來源本模型沒有收錄。'
            f'處理方式：Sens 頁業務主因表列三組案例（設定值全部 Assumed，取自 Gemini 試算與提案 G-T05，無實證，基準不採），頁末「{TITLE}」另列三案同時成立與檢驗指標。'
            f'重算結果（加權目標價，美元）：相對倍數溢價保留 50% → "&{S("rm50")}'
            f'&"；WACC 加 100bp → "&{S("w100")}&"；情境權重改 15／55／30 → "&{S("wt")}&"；三案同時成立 → "&TEXT(Sens!C{A},"0.00")&"（相對基準 "&TEXT(Sens!D{A},"+0.0%;-0.0%")'
            f'&"，評等 "&Sens!J{A}&"）。觀察：現行相對倍數（Val 0.6 節）、FY27 Q4 營收對指引與年底存貨（Recon I2）；WACC 加點在模型內沒有找到可檢驗的指標。"')
    # 註：文字刻意不寫成「加權目標價 249.59」的樣式——display_check.json 的樣式會把緊跟在「加權目標價」後的數字當成橫幅主值來核對。
    RD.cell(rr, 1).alignment = WRAP
    return f"w3_post：Sens!A{t0} 起執行長異動 {len(rows)} 列（快照 {'有' if ok else '無'}；單案數字指向 {src}）；README 風險登錄 A{rr}"
