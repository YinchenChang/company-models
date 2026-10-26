"""PR 自動獨立核對（API 版）。

讀取 PR 的文字變更（diff），請 Claude 以「未參與製作的核對者」身分檢查數字與推導，
把結果留言在 PR 上。只供參考，不阻擋合併；任何錯誤都只留言、結束碼為 0。
只用 Python 標準函式庫。
"""
import json
import os
import urllib.error
import urllib.request

REPO = os.environ["REPO"]                      # 例：YinchenChang/Tokenomics
PR = os.environ["PR_NUMBER"]
GH_TOKEN = os.environ["GH_TOKEN"]
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
MODEL = os.environ.get("REVIEW_MODEL", "claude-opus-5-5")
MAX_CHARS = int(os.environ.get("REVIEW_MAX_CHARS", "120000"))   # 送出的變更內容上限（控制費用）
PER_FILE = 30000                                                 # 單檔上限
SKIP_EXT = (".xlsx", ".xls", ".xlsm", ".png", ".jpg", ".jpeg", ".gif", ".pdf",
            ".zip", ".pkl", ".parquet", ".ico", ".woff", ".woff2")

SYSTEM = """你是「未參與製作的獨立核對者」，核對 Andy 研究 repo 的 PR。用繁體中文（台灣用語）回覆，正式、客觀。
只根據提供的變更內容判斷，不要臆測沒看到的檔案。重點依序：
1. 數字：報告、說明文字、CSV、程式常數之間的數字是否一致；能驗算的算式實際驗算（寫出算式）。
2. 推導：比值、換算、單位、口徑（例如 IT MW 與設施 MW、FP8 與 FP4、年化與全年）是否一致。
3. 來源標記：[Verified]／[Interested-party]／[Analogy]／[Assumed]／[Derived] 是否與內容相符；Assumed 或 Analogy 是否誤寫成精確點估計。
4. 文件與程式不一致、過時說明、版本號或日期前後矛盾。
輸出格式（Markdown）：
**結論**：一句話（未發現問題／發現 N 項需確認）。
**發現**：表格，欄位為 嚴重度（高／中／低）｜位置（檔名與行或段落）｜問題｜依據或驗算。沒有就寫「無」。
**未能核對**：因內容被截斷、二進位檔（如 xlsx）或缺少上下文而無法檢查的部分。
不要讚美、不要重述 PR 內容、不要提出與核對無關的建議。"""


def gh(path):
    req = urllib.request.Request(
        f"https://api.github.com/repos/{REPO}/{path}",
        headers={"Authorization": f"Bearer {GH_TOKEN}", "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r)


def comment(body):
    req = urllib.request.Request(
        f"https://api.github.com/repos/{REPO}/issues/{PR}/comments",
        data=json.dumps({"body": body}).encode(), method="POST",
        headers={"Authorization": f"Bearer {GH_TOKEN}", "Accept": "application/vnd.github+json"})
    urllib.request.urlopen(req, timeout=60).read()


def collect():
    """回傳 (送出的文字, 被略過或截斷的檔案清單)。"""
    pr = gh(f"pulls/{PR}")
    files, page = [], 1
    while page <= 10:                       # 最多 1,000 個檔
        batch = gh(f"pulls/{PR}/files?per_page=100&page={page}")
        files += batch
        if len(batch) < 100:
            break
        page += 1
    parts = [f"# PR #{PR}：{pr['title']}\n\n## PR 說明\n{(pr.get('body') or '')[:8000]}\n"]
    skipped, used = [], len(parts[0])
    for f in files:
        name = f["filename"]
        if name.lower().endswith(SKIP_EXT) or "patch" not in f:
            skipped.append(f"{name}（二進位或無文字 diff）")
            continue
        patch = f["patch"]
        if len(patch) > PER_FILE:
            patch = patch[:PER_FILE]
            skipped.append(f"{name}（單檔截斷）")
        block = f"\n## 檔案：{name}（{f['status']}）\n```diff\n{patch}\n```\n"
        if used + len(block) > MAX_CHARS:
            skipped.append(f"{name}（超過總量上限未送出）")
            continue
        parts.append(block)
        used += len(block)
    return "".join(parts), skipped


def call_claude(text):
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=json.dumps({"model": MODEL, "max_tokens": 4000, "system": SYSTEM,
                         "messages": [{"role": "user", "content": text}]}).encode(),
        headers={"x-api-key": API_KEY, "anthropic-version": "2023-06-01",
                 "content-type": "application/json"})
    with urllib.request.urlopen(req, timeout=600) as r:
        return json.load(r)


def main():
    header = "### API 自動核對（參考用，非 chat 端驗收）\n"
    if not API_KEY:
        comment(header + "未讀到 ANTHROPIC_API_KEY，略過。")
        return
    try:
        text, skipped = collect()
        resp = call_claude(text)
        out = "".join(b.get("text", "") for b in resp.get("content", []))
        u = resp.get("usage", {})
        foot = (f"\n\n---\n模型 `{resp.get('model')}`｜input {u.get('input_tokens')}｜"
                f"output {u.get('output_tokens')} tokens")
        if skipped:
            foot += "\n\n未送出核對：" + "；".join(skipped[:30])
        comment(header + out + foot)
    except urllib.error.HTTPError as e:
        comment(header + f"呼叫失敗 HTTP {e.code}：{e.read()[:500].decode(errors='replace')}")
    except Exception as e:  # 任何錯誤都不讓工作流程失敗
        comment(header + f"執行錯誤：{type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
