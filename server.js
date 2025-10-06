const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const imageModule = require('docxtemplater-image-module-free');
// PERBAIKAN: Cara impor 'image-size' yang benar
const {imageSize} = require('image-size');

const app = express();
const port = 3000;
app.use(express.static('public'));
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

const uploadFields = [
    { name: 'foto_odp', maxCount: 1 }, { name: 'klem_ring', maxCount: 1 },
    { name: 'pipa', maxCount: 1 }, { name: 'closure', maxCount: 1 },
    { name: 'in_odc', maxCount: 1 }, { name: 'testcom', maxCount: 1 },
    { name: 'port1', maxCount: 1 }, { name: 'port2', maxCount: 1 },
    { name: 'port3', maxCount: 1 }, { name: 'port4', maxCount: 1 },
    { name: 'port5', maxCount: 1 }, { name: 'port6', maxCount: 1 },
    { name: 'port7', maxCount: 1 }, { name: 'port8', maxCount: 1 },
    { name: 'in_odp', maxCount: 1 }, { name: 'label_odp', maxCount: 1 },
    { name: 'termin', maxCount: 1 }, { name: 'barcode', maxCount: 1 },
    { name: 'splitter', maxCount: 1 }, { name: 'distri', maxCount: 1 },
    { name: 'feeder', maxCount: 1 }, { name: 'odc', maxCount: 1 },
    { name: 'survey', maxCount: 1 }, { name: 'survey2', maxCount: 1 },
    { name: 'branching', maxCount: 1 }, { name: 'end_to_end', maxCount: 1 },
    { name: 'mancore', maxCount: 1 }, { name: 'peta', maxCount: 1 }
];

const upload = multer({ 
    storage: multer.memoryStorage(),
    limits: {
        fieldSize: 50 * 1024 * 1024 // Batas 50MB untuk data teks (base64)
    }
});

const documentHandler = (req, res, templateName, outputPrefix) => {
    const uploadedFilePaths = [];
    if (req.files) {
        for (const key in req.files) {
            req.files[key].forEach(file => uploadedFilePaths.push(file.path));
        }
    }

    try {
        const templatePath = path.resolve(__dirname, 'templates', templateName);
        if (!fs.existsSync(templatePath)) {
            return res.status(500).send(`Error: Template ${templateName} tidak ditemukan.`);
        }
        
        const data = { 'NAMA LOP': req.body.nama_lop || '' };

        // --- ### PENAMBAHAN FITUR: LOGIKA PERHITUNGAN BoQ ### ---
        if (outputPrefix === 'BAUT') {
            const boqItemCount = 9; // Sesuaikan jika jumlah item BoQ berubah
            for (let i = 1; i <= boqItemCount; i++) {
                const kontrak = parseInt(req.body[`boq_${i}_kontrak`] || '0', 10);
                const aktual = parseInt(req.body[`boq_${i}_aktual`] || '0', 10);
                const selisih = aktual - kontrak;
                const tambah = selisih > 0 ? selisih : 0;
                const kurang = selisih < 0 ? Math.abs(selisih) : 0;

                data[`boq_${i}_kontrak`] = kontrak;
                data[`boq_${i}_aktual`] = aktual;
                data[`boq_${i}_tambah`] = tambah;
                data[`boq_${i}_kurang`] = kurang;
            }
        }
        // --- ### AKHIR PENAMBAHAN FITUR ### ---

        uploadFields.forEach(field => {
            const placeholderName = field.name;
            if (req.files && req.files[placeholderName] && req.files[placeholderName][0]) {
                data[placeholderName] = req.files[placeholderName][0].path;
            } else if (req.body[placeholderName]) {
                data[placeholderName] = req.body[placeholderName];
            } else {
                data[placeholderName] = false;
            }
        });

        if (data['label_odp'] && !data['foto_odp']) data['foto_odp'] = data['label_odp'];
        if (data['foto_odp'] && !data['label_odp']) data['label_odp'] = data['foto_odp'];
        
        const content = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true, linebreaks: true,
            modules: [new imageModule({
                centered: false,
                getImage: (tag) => {
                    if (!tag) return null;
                    if (typeof tag === 'string' && tag.startsWith('data:image')) {
                        return Buffer.from(tag.split(',')[1], 'base64');
                    }
                    if (fs.existsSync(tag)) return fs.readFileSync(tag);
                    return null;
                },
                getSize: (imgBuffer, tag) => {
                    if (!imgBuffer) return [150, 150];
                    const dimensions = imageSize(imgBuffer); // Sekarang akan berfungsi
                    if (tag === 'end_to_end') {
                        const targetWidth = 672;
                        const ratio = targetWidth / dimensions.width;
                        return [targetWidth, Math.round(dimensions.height * ratio)];
                    }
                    if (tag === 'mancore') {
                        const targetWidth = 1008;
                        const ratio = targetWidth / dimensions.width;
                        return [targetWidth, Math.round(dimensions.height * ratio)];
                    }
                    if (tag === 'peta') {
                        const targetWidth = 787;
                        const ratio = targetWidth / dimensions.width;
                        return [targetWidth, Math.round(dimensions.height * ratio)];
                    }
                    return [dimensions.width, dimensions.height];
                }
            })],
        });

        doc.render(data);

        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        const sanitized = (req.body.nama_lop || 'Generated').replace(/\s+/g, '_');
        const fileName = `${outputPrefix} ${sanitized}.docx`;
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buf);

    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Terjadi error saat membuat dokumen. Pastikan template sudah benar.');
    } finally {
        uploadedFilePaths.forEach(filePath => {
            fs.unlink(filePath, (err) => {
                if (err) console.error(`Gagal menghapus file sementara: ${filePath}`, err);
            });
        });
        console.log("File-file sementara telah dibersihkan.");
    }
};

app.post('/generate-bact', upload.fields(uploadFields), (req, res) => {
    documentHandler(req, res, 'BACT_Template.docx', 'BACT');
});

app.post('/generate-baut', upload.fields(uploadFields), (req, res) => {
    documentHandler(req, res, 'BAUT_Template.docx', 'BAUT');
});

app.listen(port, () => {
    console.log(`🚀 Server berjalan di http://localhost:${port}`);
});
