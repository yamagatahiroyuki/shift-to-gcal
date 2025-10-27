import argparse
import re
import csv
import os
import shutil
import subprocess
import tempfile
from datetime import datetime
from typing import List, Tuple, Optional

try:
    import PyPDF2
except Exception:
    PyPDF2 = None

_Z2H_TABLE = str.maketrans({
    "０":"0","１":"1","２":"2","３":"3","４":"4",
    "５":"5","６":"6","７":"7","８":"8","９":"9",
    "：":":",
    "－":"-","ー":"-",
    "〜":"~","～":"~",
    "　":" "
})

def normalize(s: str) -> str:
    s = s.translate(_Z2H_TABLE)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def extract_text_from_pdf(pdf_path: str) -> str:
    text = ""
    if PyPDF2 is not None and pdf_path and os.path.exists(pdf_path):
        try:
            with open(pdf_path, "rb") as f:
                rdr = PyPDF2.PdfReader(f)
                text = "\n".join((p.extract_text() or "") for p in rdr.pages)
        except Exception:
            text = ""
    if text.strip():
        return text
    if shutil.which("pdftotext") and pdf_path and os.path.exists(pdf_path):
        with tempfile.TemporaryDirectory() as td:
            out_txt = os.path.join(td, "out.txt")
            try:
                subprocess.run(["pdftotext", "-layout", pdf_path, out_txt],
                               check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                with open(out_txt, "r", encoding="utf-8", errors="ignore") as fr:
                    return fr.read()
            except Exception:
                return ""
    return ""

def parse_shifts_from_text(text: str, target_name: str) -> List[Tuple[int, str, str]]:
    lines = [normalize(ln) for ln in text.splitlines() if normalize(ln)]
    results = []
    i = 0
    while i < len(lines) - 2:
        # 1) 日付の行: "1 2 3 4 5" / "6 7 8 9 10 11 12" など
        if re.fullmatch(r"(?:[1-9]|[12]\d|3[01])( (?:[1-9]|[12]\d|3[01]))+", lines[i]):
            days = [int(x) for x in lines[i].split()]
            names = lines[i+1].split()
            times = lines[i+2].split()
            n = min(len(days), len(names), len(times))
            for idx in range(n):
                if names[idx] == target_name:
                    m = re.match(r"(\d{1,2}[:：]\d{2})[~〜\-ー－](\d{1,2}[:：]\d{2})", times[idx])
                    if m:
                        start = m.group(1).replace("：", ":")
                        end   = m.group(2).replace("：", ":")
                        results.append((days[idx], start, end))
            i += 3
        else:
            i += 1
    # 重複除去
    uniq, seen = [], set()
    for d,s,e in results:
        if (d,s,e) not in seen:
            seen.add((d,s,e))
            uniq.append((d,s,e))
    return uniq


def write_gcal_min_csv(rows: List[Tuple[int, str, str]], year: int, month: int, subject: str, out_path: str):
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Subject", "Start Date", "Start Time"])
        for d, start, _ in sorted(rows, key=lambda x: x[0]):
            start_date = f"{year:04d}/{month:02d}/{d:02d}"
            w.writerow([subject, start_date, start])

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="入力PDF (OCR推奨)")
    ap.add_argument("--text", help="TXTファイルパス（Acrobatやpdftotextの書き出し）")
    ap.add_argument("--name", required=True, help="抽出対象の氏名")
    ap.add_argument("--subject", required=True, help="CSVのSubject欄")
    ap.add_argument("--year", type=int, required=True, help="年(YYYY)")
    ap.add_argument("--month", type=int, required=True, help="月(MM)")
    ap.add_argument("--output", required=True, help="出力CSVパス")
    args = ap.parse_args()

    text = ""
    if args.text and os.path.exists(args.text):
        with open(args.text, "r", encoding="utf-8", errors="ignore") as fr:
            text = fr.read()
    elif args.input and os.path.exists(args.input):
        text = extract_text_from_pdf(args.input)

    if not text.strip():
        raise SystemExit("テキストが取得できませんでした。--text にTXTを指定するか、--input にPDFを指定してください。")

    rows = parse_shifts_from_text(text, args.name)
    if not rows:
        print("警告: 抽出0件。氏名の表記ゆれ（山縣/山県、スペース有無、全角/半角）を確認してください。")
    write_gcal_min_csv(rows, args.year, args.month, args.subject, args.output)
    print(f"出力: {args.output}（{len(rows)}件）")

if __name__ == "__main__":
    main()