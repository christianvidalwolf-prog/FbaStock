import os
import shutil
from datetime import datetime


WORK_DIR = "/Users/christianvidalwolf/Stock"
CONFIG_FILE = os.path.join(WORK_DIR, "stock_backup_dir.txt")

FILES_TO_BACKUP = [
    ("STOCK AMZ.txt", "STOCK AMZ {date}.txt"),
    ("Stock Web.csv", "Stock Web {date}.csv"),
    ("cdisc precio.xlsx", "cdisc precio {date}.xlsx"),
]


def read_configured_dir():
    for env_name in ("STOCK_BACKUP_DIR", "DROPBOX_STOCK_BACKUP_DIR"):
        value = os.environ.get(env_name)
        if value:
            return os.path.expanduser(value.strip())

    if not os.path.exists(CONFIG_FILE):
        return None

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#"):
                return os.path.expanduser(line)

    return None


def main():
    print(f"--- Backing up stock files @ {datetime.now()} ---")
    backup_dir = read_configured_dir()

    if not backup_dir:
        print(
            "Backup skipped: set STOCK_BACKUP_DIR or add the local Dropbox folder "
            f"path to {CONFIG_FILE}."
        )
        return 1

    if backup_dir.startswith("http://") or backup_dir.startswith("https://"):
        print(
            "Backup skipped: Dropbox web links cannot receive file copies directly. "
            "Use the local synced Dropbox folder path instead."
        )
        return 1

    if not os.path.isdir(backup_dir):
        print(f"Backup skipped: destination folder does not exist: {backup_dir}")
        return 1

    date_str = datetime.now().strftime("%Y%m%d")
    copied = 0

    for source_name, dest_template in FILES_TO_BACKUP:
        source_path = os.path.join(WORK_DIR, source_name)
        if not os.path.exists(source_path):
            print(f"Missing source, skipped: {source_path}")
            continue

        dest_name = dest_template.format(date=date_str)
        dest_path = os.path.join(backup_dir, dest_name)
        shutil.copy2(source_path, dest_path)
        copied += 1
        print(f"Copied {source_name} -> {dest_path}")

    print(f"Backup finished: {copied}/{len(FILES_TO_BACKUP)} files copied.")
    return 0 if copied == len(FILES_TO_BACKUP) else 1


if __name__ == "__main__":
    raise SystemExit(main())
