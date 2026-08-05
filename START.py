import os, subprocess, sys
os.chdir(os.path.dirname(os.path.abspath(__file__)))
subprocess.run([sys.executable, "update_tracker.py"])
input("\nDone. Press Enter to close...")
