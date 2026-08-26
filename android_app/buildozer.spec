[app]

# (str) Tiêu đề ứng dụng hiển thị trên điện thoại
title = Quan Ly Thu Muc File

# (str) Gói: package.name phải là chữ thường, không dấu, không cách.
package.name = quanlythumucfile

# (str) Domain đảo ngược, ví dụ org.example hoặc com.tenban
package.domain = org.example

# (str) Source directory (thư mục chứa mã nguồn)
source.dir = .

# (list) Include extensions khi đóng gói
source.include_exts = py,kv,png,jpg,jpeg,gif,json,txt,db,sqlite

# (list) Exclude (loại trừ) các file/kết quả build không cần thiết
source.exclude_exts = spec,pyc,pyo
source.exclude_dirs = bin,obj,build,.buildozer,.git,__pycache__
source.exclude_patterns = .buildozer/*,bin/*

# (str) Phiên bản ứng dụng
version = 1.0

# (str) Tên file .py khởi chạy (entry point)
main.filename = main.py

# (str) Yêu cầu thư viện. KHÔNG thêm pywin32 (chỉ dành cho Windows).
# Lưu ý: 'android' và 'jnius' được python-for-android tự động cung cấp khi build Android.
requirements = python3,kivy==2.3.0,openpyxl

# (str) Không dùng Cython/NDK đặc biệt
# (bool) Tối ưu hóa (có thể tắt nếu gặp lỗi)
android.optimize_python = False

# (bool) Giao diện xoay / toàn màn hình
orientation = portrait
fullscreen = 0

# (list) Quyền Android. Cần truy cập bộ nhớ để quét thư mục.
# MANAGE_EXTERNAL_STORAGE cần thiết trên Android 11+ (API 30+) để duyệt toàn bộ bộ nhớ.
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# (int) Target API / Min API
android.api = 33
android.minapi = 21

# (int) Android NDK version
android.ndk = 23b

# (int) Android SDK version
android.sdk = 33

# (str) Android NDK download (để mặc định)
# android.ndk_path =

# (bool) Tự động chấp nhận SDK license
android.accept_sdk_license = True

# (str) Kiến trúc CPU (arm64-v8a phổ biến nhất trên điện thoại hiện nay)
android.arch = arm64-v8a

# (bool) Giữ màn hình luôn sáng
android.wakelock = False

# (list) Java classpath (mặc định)
# android.add_jars =

# (list) Thêm quyền truy cập (giống android.permissions)
# android.add_permissions =

[buildozer]

# (int) Log level (1 = error only, 2 = warning, 3 = info, 4 = debug)
log_level = 2

# (int) Màu sắc terminal (1 = có)
terminal_colored = 1

# (list) Danh sách target được hỗ trợ
targets = android,ios

# (bool) Phiên bản ổn định
warn_on_root = 0

# (str) Đường dẫn lưu APK (mặc định bin/)
# build_dir = .buildozer
