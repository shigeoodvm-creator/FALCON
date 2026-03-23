"""farm.dbのパスを確認するテストスクリプト"""
from pathlib import Path

farm_path = Path(r"C:\FARMS\デモファーム")
db_path = farm_path / "farm.db"

print(f"Farm path: {farm_path}")
print(f"Farm path exists: {farm_path.exists()}")
print(f"DB path: {db_path}")
print(f"DB path exists: {db_path.exists()}")

if farm_path.exists():
    print("\nFiles in farm folder:")
    for f in farm_path.glob("*"):
        print(f"  {f.name}")










