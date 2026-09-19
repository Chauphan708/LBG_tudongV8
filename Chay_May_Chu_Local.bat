@echo off
chcp 65001 >nul
title Khởi Chạy Ứng Dụng Local - Lịch Báo Giảng & KHDH Bản V8
echo ===================================================================
echo 🚀 ĐANG KHỞI CHẠY PHẦN MỀM LỊCH BÁO GIẢNG TIỂU HỌC BẢN V8
echo 🌐 Địa chỉ truy cập Offline/Local: http://localhost:8080
echo ===================================================================
timeout /t 1 >nul
start http://localhost:8080
node local_server.js
pause
