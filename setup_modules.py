#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script để tạo cấu trúc module js/ cho scripts.js
"""
import os
import sys

BASE_DIR = r"d:\tải xuống 2\KHO MB\sửa lại đăng nhập"
JS_DIR = os.path.join(BASE_DIR, "js")
MODULES_DIR = os.path.join(JS_DIR, "modules")

# Tạo thư mục
os.makedirs(MODULES_DIR, exist_ok=True)
print(f"Created {JS_DIR}")
print(f"Created {MODULES_DIR}")
print("Done!")
