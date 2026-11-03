# -*- coding: utf-8 -*-
"""
ساخت فایل اجرایی standalone (.exe) از داشبورد ماشین ویژن
با PyInstaller — فایل خروجی روی سیستم‌های بدون پایتون قابل اجرا است.

اجرای این فایل:
    python build_exe.py
"""
import sys
import os
import PyInstaller.__main__

BASE = os.path.dirname(os.path.abspath(__file__))

pyinstaller_args = [
    os.path.join(BASE, "app.py"),
    "--name=Mashin-Visen_Dashboard",
    "--onefile",
    "--windowed",
    "--add-data", f"{os.path.join(BASE, 'templates')}{os.pathsep}templates",
    "--add-data", f"{os.path.join(BASE, 'static')}{os.pathsep}static",
    "--hidden-import=flask",
    "--hidden-import=openpyxl",
    "--hidden-import=pandas",
    "--hidden-import=jdatetime",
    "--hidden-import=jalali_core",
    "--clean",
    "--noconfirm",
]
if sys.platform == "win32" and os.path.exists(os.path.join(BASE, "icon.ico")):
    pyinstaller_args.insert(4, "--icon=icon.ico")

PyInstaller.__main__.run(pyinstaller_args)
