"""Cross-platform build: python build.py  ->  dist/ExcelToJSON(.app|.exe|/)"""
import platform
import PyInstaller.__main__ as pyi

args = ["main.py", "--noconfirm", "--windowed", "--name", "ExcelToJSON"]
if platform.system() == "Windows":
    args.append("--onefile")  # single .exe
pyi.run(args)
