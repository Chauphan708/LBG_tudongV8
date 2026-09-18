const fs = require('fs');
const path = require('path');
const JSZip = require('./lib/jszip.min.js');

const srcDir = path.resolve(__dirname);
const targetDir = path.resolve(__dirname, '..', 'Phan_Mem_Lich_Bao_Giang_Lop_5_V7');
const targetZip = path.resolve(__dirname, '..', 'Phan_Mem_Lich_Bao_Giang_Lop_5_V7.zip');

console.log("Nguồn:", srcDir);
console.log("Đích thư mục:", targetDir);
console.log("Đích file zip:", targetZip);

// 1. Xóa thư mục đích cũ nếu có
if (fs.existsSync(targetDir)) {
    console.log("Đang xóa thư mục V7 cũ...");
    fs.rmSync(targetDir, { recursive: true, force: true });
}
if (fs.existsSync(targetZip)) {
    console.log("Đang xóa file zip V7 cũ...");
    fs.unlinkSync(targetZip);
}

fs.mkdirSync(targetDir, { recursive: true });

// 2. Danh sách thư mục & file cần copy
const itemsToCopy = [
    'index.html',
    'README.md',
    'Huong_Dan_Su_Dung_V7.txt',
    'Mo_Phan_Mem.bat',
    'Chay_May_Chu_Local.bat',
    'server.js',
    'assets',
    'css',
    'data',
    'js',
    'lib'
];

function copyRecursive(src, dest) {
    const stats = fs.statSync(src);
    if (stats.isDirectory()) {
        fs.mkdirSync(dest, { recursive: true });
        const entries = fs.readdirSync(src);
        for (const entry of entries) {
            copyRecursive(path.join(src, entry), path.join(dest, entry));
        }
    } else {
        fs.copyFileSync(src, dest);
    }
}

for (const item of itemsToCopy) {
    const s = path.join(srcDir, item);
    const d = path.join(targetDir, item);
    if (fs.existsSync(s)) {
        console.log(`Đang sao chép: ${item}...`);
        copyRecursive(s, d);
    } else {
        console.warn(`Cảnh báo: Không tìm thấy ${item}`);
    }
}

console.log("Sao chép thư mục hoàn tất! Bắt đầu tạo file ZIP chuẩn Open-Standard (Forward Slash)...");

const zip = new JSZip();

function addFolderToZip(folderPath, zipFolder) {
    const entries = fs.readdirSync(folderPath);
    for (const entry of entries) {
        const fullPath = path.join(folderPath, entry);
        const stat = fs.statSync(fullPath);
        if (stat.isDirectory()) {
            const nextZipFolder = zipFolder.folder(entry);
            addFolderToZip(fullPath, nextZipFolder);
        } else {
            const data = fs.readFileSync(fullPath);
            zipFolder.file(entry, data);
        }
    }
}

addFolderToZip(targetDir, zip);

zip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 }
}).then(buffer => {
    fs.writeFileSync(targetZip, buffer);
    const sizeMB = (buffer.length / (1024 * 1024)).toFixed(2);
    console.log(`\n🎉 ĐÓNG GÓI THÀNH CÔNG BẢN V7!`);
    console.log(`📁 Thư mục: ${targetDir}`);
    console.log(`📦 Tệp ZIP: ${targetZip} (${sizeMB} MB)`);
}).catch(err => {
    console.error("Lỗi khi nén ZIP:", err);
});
