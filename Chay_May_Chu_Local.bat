@echo off
chcp 65001 >nul
title Khởi Chạy Ứng Dụng Local - Lịch Báo Giảng Lớp 5 Bản V4
echo ===================================================================
echo 🚀 ĐANG KHỞI CHẠY PHẦN MỀM LỊCH BÁO GIẢNG LỚP 5 BẢN V4
echo 🌐 Địa chỉ truy cập Offline/Local: http://localhost:8080
echo ===================================================================
timeout /t 1 >nul
start http://localhost:8080
node server.js
pause
