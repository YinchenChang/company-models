# tools/sens_coverage/check_sens_coverage.py 的單元測試（暫存小活頁簿）。
# 執行：python3 -m pytest -q tools/sens_coverage/test_check_sens_coverage.py  或  python3 tools/sens_coverage/test_check_sens_coverage.py
import os, subprocess, sys, tempfile
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import check_sens_coverage as C  # noqa: E402

HE = [("he_h", "IF_HoldEcon｜H100", 6.8, 9.7, 16.7), ("he_b2", "IF_HoldEcon｜GB200", 6.6, 8.8, 14.3),
      ("he_b3", "IF_HoldEcon｜GB300", 8.4, 10.9, 18.5), ("he_vr", "IF_HoldEcon｜VR200", 9.2, 12.2, 19.8)]


def make(path, in_rows, sens_rows, a2=None, in_sheet=True, three_col=True, sens=True):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    if in_sheet:
        ws = wb.create_sheet("In")
        ws.append(["In"]); ws.append([None])
        ws.append(["鍵", "項目", "單位"] + (["敏感度低", "基準", "敏感度高"] if three_col else ["值"]) + ["標記", "來源", "說明"])
        for k, lab, d, e, f, *note in in_rows:
            ws.append([k, lab, "x", d, e, f, "Judgment", "src", note[0] if note else None])
        ws.append([None]); ws.append(["情境驅動對照（純展示）"]); ws.append(["鍵", "項目", "單位", "情境 A", "情境 B", "情境 C"])
        ws.append(["zz_A", "不該被讀", "x", 1, 2, 3])  # 本節以下不列入
    if sens:
        s = wb.create_sheet("Sens")
        s.append(["Sens — 測試"]); s.append([a2])
        s.append(["驅動", "設定", "方向", "加權目標價", "相對基準", "鍵", "說明"])
        for drv, key, note in sens_rows:
            s.append([drv, "低／高", "低", 100, 0, key, note])
    wb.save(path)
    return path


BASE_IN = [("tax", "有效稅率", 0.13, 0.16, 0.19), ("same", "低高同基準", 1, 1, 1), ("only_hi", "只有高", None, 0, 1)] + HE
BASE_SENS = [("有效稅率：低", "tax", "重算值"), ("只有高：高", "only_hi", "重算值"), ("IF_HoldEcon 四世代同時：低", "he_all", "重算值")]


def tmp(name):
    return os.path.join(tempfile.mkdtemp(prefix="senscov_"), name)


def test_pass_with_he_all():
    r = C.check(make(tmp("ok.xlsx"), BASE_IN, BASE_SENS))
    assert r["applicable"] and not r["missing"] and r["keys"] == 6 and "IF_HoldEcon" in r["group_cov"]


def test_missing_he_fails_even_if_exempted():
    sens = [x for x in BASE_SENS if x[1] != "he_all"]
    a2 = "he_h、he_b2、he_b3、he_vr 不列：測試用豁免（群組不得豁免）。"
    p = make(tmp("nohe.xlsx"), BASE_IN, sens, a2=a2)
    r = C.check(p)
    assert sorted(m[0] for m in r["missing"]) == ["he_b2", "he_b3", "he_h", "he_vr"]
    assert all("不得豁免" in m[5] for m in r["missing"])
    assert C.main([p]) == 1


def test_single_he_member_case_not_enough():
    sens = [x for x in BASE_SENS if x[1] != "he_all"] + [("GB300 持有成本：高", "he_b3", "重算值")]
    r = C.check(make(tmp("he1.xlsx"), BASE_IN, sens))
    assert {m[0] for m in r["missing"]} == {"he_h", "he_b2", "he_b3", "he_vr"}


def test_he_group_by_text():
    sens = [x for x in BASE_SENS if x[1] != "he_all"] + [("Tokenomics IF_HoldEcon 四世代同時：低", "", "")]
    r = C.check(make(tmp("hetxt.xlsx"), BASE_IN, sens))
    assert not r["missing"]


def test_exemption_in_a2_passes():
    ins = BASE_IN + [("inc_m", "增量 EBITDA 率", 0.75, 0.85, 0.95)]
    a2 = "相對基準＝…。inc_m（增量 EBITDA 率）不列：A、C 對稱，對加權目標價恆為 0。"
    p = make(tmp("ex.xlsx"), ins, BASE_SENS, a2=a2)
    r = C.check(p)
    assert not r["missing"] and r["exempt"][0][0] == "inc_m"
    assert C.main([p]) == 0


def test_exemption_in_note_column_passes():
    ins = BASE_IN + [("gl_rl", "RL 營收長期年增", -0.05, 0.05, 0.15)]
    sens = BASE_SENS + [("（無）", "", "gl_rl 豁免：重算變動小於 0.001%")]
    assert not C.check(make(tmp("ex2.xlsx"), ins, sens))["missing"]


def test_missing_plain_key_fails_and_mention_without_exempt_word_is_not_exempt():
    ins = BASE_IN + [("wacc_adj", "WACC 調整", -0.01, 0, 0.01)]
    p = make(tmp("miss.xlsx"), ins, BASE_SENS, a2="wacc_adj 見 Val 頁")  # 提到鍵但沒有豁免字
    r = C.check(p)
    assert [m[0] for m in r["missing"]] == ["wacc_adj"]


def test_key_mentioned_in_case_text_counts():
    ins = BASE_IN + [("cip_ai", "在建工程中 AI 占比", 0.8, 0.9, 1)]
    sens = BASE_SENS + [("在建工程不另增 MW（cip_ai＝0）：低", "other_key", "")]
    assert not C.check(make(tmp("txt.xlsx"), ins, sens))["missing"]


def test_not_model_note_skipped():
    ins = BASE_IN + [("cons_tp", "市場共識目標價", 75, 245, 800, "低／高＝最低／最高（只作對照，不進模型）")]
    r = C.check(make(tmp("nm.xlsx"), ins, BASE_SENS))
    assert not r["missing"] and r["not_model"] == ["cons_tp"]


def test_no_in_sheet_not_applicable():
    p = make(tmp("noin.xlsx"), [], BASE_SENS, in_sheet=False)
    assert C.check(p)["applicable"] is False
    assert C.main([p]) == 0


def test_in_without_three_cols_not_applicable():
    p = make(tmp("old.xlsx"), [("x", "舊模板", 1, 2, 3)], [], three_col=False)
    assert C.check(p)["applicable"] is False


def test_applicable_without_sens_fails():
    p = make(tmp("nosens.xlsx"), BASE_IN, [], sens=False)
    r = C.check(p)
    assert r["error"] and len(r["missing"]) == 6
    assert C.main([p]) == 1


def test_cli_runs():
    p = make(tmp("cli.xlsx"), BASE_IN, BASE_SENS)
    out = subprocess.run([sys.executable, os.path.join(HERE, "check_sens_coverage.py"), p], capture_output=True, text=True)
    assert out.returncode == 0 and "通過" in out.stdout


if __name__ == "__main__":
    fs = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for f in fs:
        f()
    print(f"{len(fs)} 項測試通過")
