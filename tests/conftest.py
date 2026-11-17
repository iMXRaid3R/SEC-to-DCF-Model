import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))                  # import the edgar package
sys.path.insert(0, str(Path(__file__).parent))  # import xl_com
