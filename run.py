import os
import subprocess
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

TASKS = [
    ("1/2  位置精度", "position.py"),
    ("2/2  可靠性", "reliability.py"),
]


if __name__ == "__main__":
    for label, script in TASKS:
        print("=" * 50)
        print(label)
        print("=" * 50)
        # 用当前解释器执行，避免调用到 PATH 中的其他 Python
        ret = subprocess.run([sys.executable, os.path.join(BASE_DIR, script)])
        if ret.returncode != 0:
            print(f"[ERROR] {script} 退出码 {ret.returncode}")
            sys.exit(ret.returncode)
        print()

    print("=" * 50)
    print("全部完成")
    print("=" * 50)
