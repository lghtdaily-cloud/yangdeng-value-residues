# -*- coding: utf-8 -*-
import subprocess, sys, json
sys.stdout.reconfigure(encoding='utf-8')

ns = {}
src = open(r"C:\Users\美人鱼\WorkBuddy\2026-09-18-14-33-33\改写字数校验.py", encoding="utf-8").read()
exec(src.split("for n, t in")[0], ns)
t42, t51, t52 = ns["t42"], ns["t51"], ns["t52"]

SKILL = r"C:\Users\美人鱼\WorkBuddy\resources\app.asar.unpacked\resources\plugins\workbuddy-builtin\skills\tencent-local-office-edit"
FID = "C:/Users/美人鱼/OneDrive/Desktop/羊磴磴队实践创新报告.docx"

def call(tool, **kw):
    args = [sys.executable, "edsdk.py", "call", tool, "file_id=" + FID]
    for k, v in kw.items():
        args.append("%s=%s" % (k, v if isinstance(v, str) else json.dumps(v, ensure_ascii=False)))
    r = subprocess.run(args, cwd=SKILL, capture_output=True, text=True, encoding="utf-8", errors="replace")
    return r

# 每个单元格：先删除多余正文段（从后往前），再替换首段正文
jobs = [
    ("5-2", [7973, 7816, 7684, 7648, 7432], (7419, 7431), t52[8:]),
    ("5-1", [7146, 6961, 6766], (6602, 6765), t51[9:]),
    ("4-2", [6330, 6181, 6006, 5842], (5808, 5841), t42[12:]),
]

for name, dels, (b, e), txt in jobs:
    for idx in dels:
        r = call("doc_delete_paragraph", idx=str(idx))
        print(name, "del", idx, "rc=", r.returncode, (r.stdout or r.stderr)[:120].replace("\n", " "))
    r = call("doc_replace_text", text=txt, ranges=[{"begin": b, "end": e}])
    print(name, "replace", "rc=", r.returncode, (r.stdout or r.stderr)[:160].replace("\n", " "))
