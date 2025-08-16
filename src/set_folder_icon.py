# -*- coding: utf-8 -*-
import os
import re
import sys
import time
import subprocess
import ctypes

BASE_DIR = r"C:\Users\ken1r\projects\work"
ICON_PATH = r"C:\Users\ken1r\projects\icon_assets\@ico\Benjigarner-Rise-Folder-Summer-generic.256.ico"

INVALID_CHARS_PATTERN = r'[<>:"/\\|?*\x00-\x1F]'
RESERVED_NAMES = {
    "CON","PRN","AUX","NUL",
    *(f"COM{i}" for i in range(1,10)),
    *(f"LPT{i}" for i in range(1,10)),
}

def is_reserved_name(name: str) -> bool:
    base = os.path.splitext(name)[0].upper()
    return base in RESERVED_NAMES

def sh_change_notify():
    SHCNE_ASSOCCHANGED = 0x08000000
    SHCNF_IDLIST = 0x0000
    try:
        ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
    except Exception:
        pass

def set_folder_icon(target_folder: str, icon_path: str):
    """
    desktop.ini を生成し、UTF-16LEで保存
    フォルダに+S属性、desktop.iniに+H+S属性を付与
    """
    desktop_ini = os.path.join(target_folder, "desktop.ini")
    ini_content = (
        "[.ShellClassInfo]\r\n"
        f"IconResource={icon_path},0\r\n"
        f"IconFile={icon_path}\r\n"
        "IconIndex=0\r\n"
        "ConfirmFileOp=0\r\n"
    )

    # UTF-16LE（BOM付き）で書き込み
    with open(desktop_ini, "w", encoding="utf-16") as f:
        f.write(ini_content)

    # desktop.ini を隠し＋システム
    subprocess.run(['attrib', '+h', '+s', desktop_ini], shell=True)

    # フォルダにSystem属性を付与（これがないと反映されない）
    subprocess.run(['attrib', '+s', target_folder], shell=True)

def main():
    if len(sys.argv) < 2:
        print("❌ フォルダ名が指定されていません。")
        time.sleep(2)
        sys.exit(1)

    folder_name = sys.argv[1].strip()
    print(f"▶ 受け取った文字列: {folder_name}")
    time.sleep(2)

    # 不正文字
    if re.search(INVALID_CHARS_PATTERN, folder_name):
        print("❌ この名前にはフォルダ名に使えない文字が含まれています。")
        sys.exit(1)

    if is_reserved_name(folder_name):
        print("❌ この名前はWindowsの予約語です。")
        sys.exit(1)

    if folder_name in {".", ".."} or len(folder_name) == 0:
        print("❌ 無効なフォルダ名です。")
        sys.exit(1)

    if not os.path.isdir(BASE_DIR):
        print(f"❌ ベースディレクトリが存在しません: {BASE_DIR}")
        sys.exit(1)

    target = os.path.join(BASE_DIR, folder_name)

    if os.path.exists(target):
        print(f"ℹ 既に存在します: {target}")
        subprocess.Popen(['explorer', target])
        sys.exit(0)

    # フォルダ構成
    os.makedirs(os.path.join(target, "src"), exist_ok=True)
    os.makedirs(os.path.join(target, "archive"), exist_ok=True)
    readme_path = os.path.join(target, "README.md")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(f"# {folder_name}\n\n- src: 作業ファイル\n- archive: 過去ファイル\n")

    # アイコン設定
    if os.path.isfile(ICON_PATH):
        set_folder_icon(target, ICON_PATH)
    else:
        print(f"⚠ アイコンが見つかりません: {ICON_PATH}")

    # 更新通知＋開く
    sh_change_notify()
    subprocess.Popen(['explorer', target])
    print("✅ 完了")

if __name__ == "__main__":
    main()
