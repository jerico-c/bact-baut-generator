const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const imageModule = require('docxtemplater-image-module-free');
const {imageSize} = require('image-size');

const app = express();
const port = process.env.PORT || 8080; 

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
        fieldSize: 50 * 1024 * 1024
    }
});

const generateDocumentHandler = (req, res, templateName, outputPrefix) => {
    try {
        const templatePath = path.resolve(__dirname, 'templates', templateName);
        if (!fs.existsSync(templatePath)) {
            return res.status(500).send(`Error: Template ${templateName} tidak ditemukan.`);
        }
        
        const data = { 'NAMA LOP': req.body.nama_lop || '' };

        // Read the template XML to detect which placeholders are used as loops
        // so we can provide arrays for those fields (e.g., {#label_odp} ... {/label_odp})
        let templateBinaryForDetect;
        let templateDocXml = '';
        try {
            templateBinaryForDetect = fs.readFileSync(templatePath, 'binary');
            const zipDetect = new PizZip(templateBinaryForDetect);
            const docFile = zipDetect.file('word/document.xml');
            if (docFile) templateDocXml = docFile.asText();
        } catch (e) {
            // ignore detection errors; fallback behavior will still work
            templateDocXml = '';
        }

        if (outputPrefix === 'BAUT') {
            const boqItemCount = 9;
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

        uploadFields.forEach(field => {
            const placeholderName = field.name;
            const isLoop = templateDocXml.includes(`{#${placeholderName}`);
            // Prefer multer-uploaded file -> convert to data URL string so the image-module
            // will call our getImage/getSize (it only calls those for non-object tag values).
            if (req.files && req.files[placeholderName] && req.files[placeholderName][0]) {
                const file = req.files[placeholderName][0];
                const dataUrl = `data:${file.mimetype};base64,${file.buffer.toString('base64')}`;
                data[placeholderName] = isLoop ? [dataUrl] : dataUrl;
            } else if (req.body[placeholderName] && typeof req.body[placeholderName] === 'string' && req.body[placeholderName].startsWith('data:image')) {
                // base64 data URL from client (keep as string so module will decode it)
                data[placeholderName] = isLoop ? [req.body[placeholderName]] : req.body[placeholderName];
            } else {
                // use null for missing images (or empty array when loop is expected)
                data[placeholderName] = isLoop ? [] : null;
            }
        });

        if (data['label_odp'] && !data['foto_odp']) data['foto_odp'] = data['label_odp'];
        if (data['foto_odp'] && !data['label_odp']) data['label_odp'] = data['foto_odp'];
        
    // reuse template binary if we already read it for detection, otherwise read now
    const content = templateBinaryForDetect || fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true, linebreaks: true,
            modules: [new imageModule({
                centered: false,
                // getImage should return a Buffer (or null). The tag parameter can be many shapes
                getImage: (tag) => {
                    // tag can be: Buffer, base64 string, array (from loops), or an object
                    if (!tag) return null;
                    // If it's already a Buffer
                    if (Buffer.isBuffer(tag)) return tag;
                    // If it's an array, try first element
                    if (Array.isArray(tag) && tag.length > 0) {
                        const first = tag[0];
                        if (Buffer.isBuffer(first)) return first;
                        if (typeof first === 'string' && first.startsWith('data:image')) {
                            try { return Buffer.from(first.split(',')[1], 'base64'); } catch (e) { return null; }
                        }
                    }
                    // If it's a data URL string
                    if (typeof tag === 'string' && tag.startsWith('data:image')) {
                        try { return Buffer.from(tag.split(',')[1], 'base64'); } catch (e) { return null; }
                    }
                    // If it's an object that may contain a buffer (e.g., multer file-like)
                    if (typeof tag === 'object') {
                        if (tag.buffer && Buffer.isBuffer(tag.buffer)) return tag.buffer;
                        if (tag.data && Buffer.isBuffer(tag.data)) return tag.data;
                    }
                    return null;
                },
                // getSize must always return [width, height]
                getSize: (imgBuffer, tag) => {
                    if (!imgBuffer) return [150, 150];
                    try {
                        const dimensions = imageSize(imgBuffer);
                        if (!dimensions || !dimensions.width || !dimensions.height) return [150, 150];
                        const pageMaxWidth = 672;
                        let targetWidth;

                        // tag may be the original tag value or a name; normalize to string when appropriate
                        const tagName = (typeof tag === 'string') ? tag : (tag && tag.name) ? tag.name : null;

                        switch (tagName) {
                            case 'end_to_end': targetWidth = 671; break;
                            case 'mancore': targetWidth = 1008; break;
                            case 'peta': targetWidth = 786; break;
                            default: targetWidth = Math.min(dimensions.width, 300); break;
                        }

                        targetWidth = Math.min(targetWidth, pageMaxWidth);
                        const ratio = targetWidth / dimensions.width;
                        return [Math.round(targetWidth), Math.round(dimensions.height * ratio)];
                    } catch (e) {
                        // If image-size fails, fall back to a safe default
                        return [150, 150];
                    }
                },
            })],
        });

        try {
            doc.render(data);
        } catch (e) {
            // Better error message for template rendering issues
            console.error('Render error:', e);
            // If the error contains properties from docxtemplater, include them
            if (e.properties && e.properties.errors) {
                console.error('Detailed errors:', e.properties.errors.map(err => err && err.properties && err.properties.explanation).join('\n'));
            }
            return res.status(500).send('Terjadi error saat merender template. Cek logs untuk detail.');
        }

        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        const sanitized = (req.body.nama_lop || 'Generated').replace(/[\s/]/g, '_');
        const fileName = `${outputPrefix} ${sanitized}.docx`;
        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.send(buf);

    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Terjadi error saat membuat dokumen. Pastikan template sudah benar.');
    } 
    // --- PERBAIKAN: Blok 'finally' dan kode pembersihan file telah dihapus ---
};

app.post('/generate-bact', upload.fields(uploadFields), (req, res) => {
    generateDocumentHandler(req, res, 'BACT_Template.docx', 'BACT');
});

app.post('/generate-baut', upload.fields(uploadFields), (req, res) => {
    generateDocumentHandler(req, res, 'BAUT_Template.docx', 'BAUT');
});

app.listen(port, () => {
    console.log(`🚀 Server berjalan di http://localhost:${port}`);
});
