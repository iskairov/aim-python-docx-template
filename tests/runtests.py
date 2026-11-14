# -*- coding: utf-8 -*-
"""
Скрипт запуска всех тестовых сценариев генерации документов в директории tests/.
"""
from __future__ import annotations

import glob
import os
import subprocess
import sys

base_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(base_dir)
output_dir = os.path.join(base_dir, "output")

os.chdir(base_dir)
tests = sorted(glob.glob("[A-Za-z]*.py"))
excludes = ["runtests.py"]

if not os.path.exists(output_dir):
    os.mkdir(output_dir)

env = os.environ.copy()
existing_pythonpath = env.get("PYTHONPATH", "")
env["PYTHONPATH"] = (
    f"{project_root}{os.pathsep}{existing_pythonpath}"
    if existing_pythonpath
    else project_root
)

for test in tests:
    if test not in excludes:
        print(f"{test} ...")
        rc = subprocess.call([sys.executable, f"./{test}"], cwd=base_dir, env=env)
        if rc != 0:
            sys.exit(rc)

print("Готово (Done).")
