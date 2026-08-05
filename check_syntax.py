import sys
print(f"Python {sys.version}")
try:
    with open("update_tracker.py", encoding="utf-8") as f:
        src = f.read()
    code = compile(src, "update_tracker.py", "exec")
    print("Syntax OK — no errors found")
except IndentationError as e:
    print(f"\nIndentationError on line {e.lineno}:")
    print(f"  {e.msg}")
    lines = src.splitlines()
    for i in range(max(0, e.lineno - 5), min(len(lines), e.lineno + 2)):
        marker = ">>>" if i + 1 == e.lineno else "   "
        print(f"  {marker} {i+1:4}: {repr(lines[i])}")
except SyntaxError as e:
    print(f"\nSyntaxError on line {e.lineno}: {e.msg}")
input("\nPress Enter to close...")
