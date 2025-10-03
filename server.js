// server.js (Final - Ekstraksi dari .docx berdasarkan urutan)

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

const uploadDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir);
}
const upload = multer({ dest: uploadDir });

// INI ADALAH BAGIAN PALING PENTING
// "Kontrak" antara template eviden dan server. 
// Urutan placeholder di sini HARUS SAMA PERSIS dengan urutan gambar di dalam file EVIDEN.docx
const placeholderOrder = [
    'port1', 'port2', 'port3', 'port4', 'port5', 'port6', 'port7', 'port8', 'in_odp',
    'port9', 'port10', 'port11', 'port12', 'port13', 'port14', 'port15', 'port16', 'input', // <-- Ini harusnya "in_odp" atau "input"? Sesuaikan dengan template BACT Anda. Saya asumsikan "input".
    'label_odp', 'termin', 'barcode',
    'whwl', 'pralon', 'valins', // <-- Tambahan dari .doc baru, sesuaikan jika perlu
    'output_ps_1_4', 'spliter', 'distri', // <-- Ganti nama agar valid
    'feeder', 'odc', 'mo',
    'join_branching_1', 'join_branching_2', 'progress_k3',
    'eviden_geledek', 'rumah_calang', 'survei_odc',
    'odp_cl' // <-- Tambahan dari .doc baru
];

app.post('/generate-bact', upload.single('eviden_docx'), (req, res) => {
    const tempImagePaths = [];
    if (req.file) tempImagePaths.push(req.file.path); // Untuk menghapus file .docx upload

    try {
        if (!req.file) {
            return res.status(400).send('Tidak ada file eviden .docx yang diunggah.');
        }

        // 1. Ekstrak Gambar dari Eviden.docx
        const evidenContent = fs.readFileSync(req.file.path);
        const evidenZip = new PizZip(evidenContent);
        
        const imageFiles = evidenZip.file(/word\/media\//);
        if (!imageFiles || imageFiles.length === 0) {
            return res.status(400).send('Tidak ada gambar yang ditemukan di dalam file .docx yang diunggah.');
        }
        
        const imageData = {};

        // Urutkan gambar berdasarkan nama file bawaan (image1, image2, dst.)
        imageFiles.sort((a, b) => {
            const numA = parseInt(a.name.match(/\d+/)[0], 10);
            const numB = parseInt(b.name.match(/\d+/)[0], 10);
            return numA - numB;
        });

        imageFiles.forEach((imgFile, index) => {
            if (index < placeholderOrder.length) {
                const placeholderName = placeholderOrder[index];
                const buffer = imgFile.asNodeBuffer();
                const extension = path.extname(imgFile.name);
                
                const tempPath = path.join(uploadDir, `${Date.now()}-${placeholderName}${extension}`);
                fs.writeFileSync(tempPath, buffer);
                
                imageData[placeholderName] = tempPath;
                tempImagePaths.push(tempPath);
            }
        });

        // 2. Isi Template BACT dengan Gambar yang Sudah Diekstrak
        const templatePath = path.join(__dirname, 'templates', 'BACT_Template.docx');
        if (!fs.existsSync(templatePath)) {
            return res.status(500).send('File template BACT_Template.docx tidak ditemukan di server.');
        }
        
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        
        const imageOpts = {
            centered: false,
            getImage: (tag) => fs.readFileSync(tag),
            getSize: () => [150, 150], // Sesuaikan ukuran default jika perlu
        };
        const imageModule = new ImageModule(imageOpts);

        const doc = new Docxtemplater(zip, { modules: [imageModule], paragraphLoop: true, linebreaks: true });

        const dataToRender = { NAMA_LOP: req.body.nama_lop || 'N/A' };
        placeholderOrder.forEach(key => {
            dataToRender[key] = imageData[key] || false;
        });

        doc.render(dataToRender);
        
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        
        const sanitizedLop = (req.body.nama_lop || 'UNTITLED').replace(/[\s/]/g, '_');
        const fileName = `BACT ${sanitizedLop}.docx`;

        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buf);

    } catch (error) {
        console.error('Error generating document:', error);
        res.status(500).send('Terjadi kesalahan saat membuat dokumen. Cek konsol server untuk detail.');
    } finally {
        // Hapus semua file sementara (docx upload dan gambar hasil ekstrak)
        tempImagePaths.forEach(p => {
            fs.unlink(p, (err) => {
                if (err) console.error(`Gagal menghapus file sementara: ${p}`, err);
            });
        });
        console.log('File sementara telah dibersihkan.');
    }
});

app.listen(PORT, () => {
    console.log(`Server berjalan di http://localhost:${PORT}`);
});