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
        
        const data = {
            'NAMA LOP': req.body.nama_lop || '',
            'NAMA LOKASI': req.body.nama_lop || '',
            'IHLD LoP ID': req.body.ihld_lop_id || '',
            'iHLD LoP ID': req.body.ihld_lop_id || '',
            'eProposal LoP ID': req.body.eprop_lop_id || '',
            'STO': req.body.sto || ''
        };

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

        
        // BOQ auto-calculation is optional. Only run if outputPrefix is 'BAUT'
        // and the client explicitly enabled the feature via form field 'use_boq_auto'.
        const useBoqAuto = String(req.body.use_boq_auto || '') === 'on' || String(req.body.use_boq_auto || '') === 'true';
        if (outputPrefix === 'BAUT' && useBoqAuto) {
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
            } else if (req.body[placeholderName] && typeof req.body[placeholderName] === 'string' && req.body[placeholderName].startsWith('data:')) {
                // base64 data URL from client (keep as string so module will decode it)
                data[placeholderName] = isLoop ? [req.body[placeholderName]] : req.body[placeholderName];
            } else {
                // use null for missing images (or empty array when loop is expected)
                data[placeholderName] = isLoop ? [] : null;
            }
        });

        if (data['label_odp'] && !data['foto_odp']) data['foto_odp'] = data['label_odp'];
        if (data['foto_odp'] && !data['label_odp']) data['label_odp'] = data['foto_odp'];

        // Pre-clean: if some image placeholders are null, remove raw tags like %tag/%%tag from XML
        try {
            const nullTags = uploadFields.map(f => f.name).filter(name => data[name] == null);
            if (nullTags.length > 0) {
                const cleanFiles = ['word/document.xml', 'word/header1.xml', 'word/footer1.xml'];
                // ensure we have binary loaded to zip for editing
                const preContent = templateBinaryForDetect || fs.readFileSync(templatePath, 'binary');
                const preZip = new PizZip(preContent);
                cleanFiles.forEach(fname => {
                    const f = preZip.file(fname);
                    if (!f) return;
                    let text = f.asText();
                    nullTags.forEach(tag => {
                        const pats = [new RegExp('%%' + tag + '\\b', 'g'), new RegExp('%' + tag + '\\b', 'g')];
                        pats.forEach(p => { text = text.replace(p, ''); });
                    });
                    preZip.file(fname, text);
                });
                // Reassign the modified binary for the actual compile step
                templateBinaryForDetect = preZip.generate({ type: 'binary' });
            }
        } catch (precleanErr) {
            console.error('[DEBUG] Pre-clean placeholders failed:', precleanErr);
        }
        
    // reuse template binary if we already read it for detection, otherwise read now
    const content = templateBinaryForDetect || fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
        let doc;
        let usedImageModule = true;
        const renderWarnings = [];
        try {
            doc = new Docxtemplater(zip, {
                paragraphLoop: true,
                linebreaks: true,
                nullGetter: () => '',
                modules: [new imageModule({
                    setParser: (placeHolderContent) => {
                        if (typeof placeHolderContent !== 'string') return null;
                        if (placeHolderContent.startsWith('%%')) {
                            return { type: 'placeholder', value: placeHolderContent.substr(2).toLowerCase(), module: 'open-xml-templating/docxtemplater-image-module', centered: true };
                        }
                        if (placeHolderContent.startsWith('%')) {
                            return { type: 'placeholder', value: placeHolderContent.substr(1).toLowerCase(), module: 'open-xml-templating/docxtemplater-image-module', centered: false };
                        }
                        return null;
                    },
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
                    // If it's a data URL string (accept any data:* to be tolerant)
                        if (typeof tag === 'string' && tag.startsWith('data:')) {
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

                        // convert centimeters to pixels using 96 DPI (common for Office)
                        const DPI = 96;
                        const PIXELS_PER_CM = DPI / 2.54; // ~37.79527559

                        // helper to robustly extract a placeholder name from tag
                        const extractTagName = (t) => {
                            if (!t) return null;
                            if (typeof t === 'string') return t;
                            if (Array.isArray(t) && t.length > 0) return extractTagName(t[0]);
                            if (typeof t === 'object') {
                                if (t.name && typeof t.name === 'string') return t.name;
                                if (t.tag && typeof t.tag === 'string') return t.tag;
                                if (t.key && typeof t.key === 'string') return t.key;
                            }
                            return null;
                        };

                        let tagName = extractTagName(tag);
                        // normalize to a searchable string form
                        const tagStr = tagName ? String(tagName).toLowerCase() : String(tag || '').toLowerCase();

                        // Default: constrain width to a reasonable maximum (in pixels)
                        const pageMaxWidth = Math.round(17.76 * PIXELS_PER_CM); // use same max as target for full-width cases
                        let targetWidth;

                        // If DEBUG is set, log tag info (helps troubleshoot tag shapes)
                        if (process.env.DEBUG) {
                            console.error('[DEBUG] getSize tag:', { rawTag: tag, extracted: tagName, tagStr });
                        }

                        // Make end_to_end exactly 17.76 cm wide when the tag name matches or includes the token
                        if (tagStr.includes('end_to_end') || tagStr.includes('endtoend') || tagStr.includes('end-to-end')) {
                            targetWidth = Math.round(17.76 * PIXELS_PER_CM); // 17.76 cm -> pixels at 96 DPI => ~672 px
                        } else {
                            switch (tagName) {
                                case 'mancore':
                                    targetWidth = 1008;
                                    break;
                                case 'peta':
                                    targetWidth = 786;
                                    break;
                                default:
                                    targetWidth = Math.min(dimensions.width, 300);
                                    break;
                            }
                        }

                        // Ensure we don't exceed the page max width
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
        } catch (compileErr) {
            console.error('Docxtemplater compile error with image module:', compileErr && compileErr.properties ? compileErr.properties : compileErr);
            // Try to auto-remove problematic raw tags reported by the compile error and retry
            try {
                const errors = (compileErr && compileErr.properties && compileErr.properties.errors) || [];
                const badTags = Array.from(new Set(errors.map(err => (err && err.xtag) || (err && err.properties && err.properties.xtag)).filter(Boolean)));
                if (badTags.length > 0) {
                    renderWarnings.push(`Removed problematic image tags: ${badTags.join(', ')}`);
                    // Attempt to remove raw occurrences of these tags from document.xml and any other xml files
                    badTags.forEach(tag => {
                        ['word/document.xml', 'word/footer1.xml', 'word/header1.xml'].forEach(fname => {
                            const f = zip.file(fname);
                            if (f) {
                                let text = f.asText();
                                // Remove occurrences of %tag and %%tag and the raw %tag patterns
                                const patterns = [
                                    new RegExp('%' + tag, 'g'),
                                    new RegExp('%%' + tag, 'g'),
                                    new RegExp('%' + tag + '\\b', 'g')
                                ];
                                patterns.forEach(p => { text = text.replace(p, ''); });
                                zip.file(fname, text);
                            }
                        });
                    });
                    // Retry compilation with modified zip
                    try {
                        doc = new Docxtemplater(zip, {
                            paragraphLoop: true,
                            linebreaks: true,
                            nullGetter: () => '',
                            modules: [new imageModule({ centered: false,
                                setParser: (placeHolderContent) => {
                                    if (typeof placeHolderContent !== 'string') return null;
                                    if (placeHolderContent.startsWith('%%')) {
                                        return { type: 'placeholder', value: placeHolderContent.substr(2).toLowerCase(), module: 'open-xml-templating/docxtemplater-image-module', centered: true };
                                    }
                                    if (placeHolderContent.startsWith('%')) {
                                        return { type: 'placeholder', value: placeHolderContent.substr(1).toLowerCase(), module: 'open-xml-templating/docxtemplater-image-module', centered: false };
                                    }
                                    return null;
                                },
                                getImage: (tag) => {
                                    if (!tag) return null;
                                    if (Buffer.isBuffer(tag)) return tag;
                                    if (Array.isArray(tag) && tag.length > 0) {
                                        const first = tag[0];
                                        if (Buffer.isBuffer(first)) return first;
                                        if (typeof first === 'string' && first.startsWith('data:image')) {
                                            try { return Buffer.from(first.split(',')[1], 'base64'); } catch (e) { return null; }
                                        }
                                    }
                                    if (typeof tag === 'string' && tag.startsWith('data:')) {
                                        try { return Buffer.from(tag.split(',')[1], 'base64'); } catch (e) { return null; }
                                    }
                                    if (typeof tag === 'object') {
                                        if (tag.buffer && Buffer.isBuffer(tag.buffer)) return tag.buffer;
                                        if (tag.data && Buffer.isBuffer(tag.data)) return tag.data;
                                    }
                                    return null;
                                },
                                getSize: (imgBuffer, tag) => {
                                    if (!imgBuffer) return [150,150];
                                    try {
                                        const dimensions = imageSize(imgBuffer) || { width: 150, height: 150 };
                                        const DPI = 96; const PIXELS_PER_CM = DPI / 2.54;
                                        const extractTagName = (t) => { if (!t) return null; if (typeof t === 'string') return t; if (Array.isArray(t) && t.length>0) return extractTagName(t[0]); if (typeof t === 'object') { if (t.name) return t.name; if (t.tag) return t.tag; if (t.key) return t.key;} return null; };
                                        const tagName = extractTagName(tag);
                                        const tagStr = tagName ? String(tagName).toLowerCase() : String(tag || '').toLowerCase();
                                        const pageMaxWidth = Math.round(17.76 * PIXELS_PER_CM);
                                        let targetWidth = Math.min(dimensions.width, 300);
                                        if (tagStr.includes('end_to_end') || tagStr.includes('endtoend') || tagStr.includes('end-to-end')) targetWidth = Math.round(17.76 * PIXELS_PER_CM);
                                        if (tagName === 'mancore') targetWidth = 1008;
                                        if (tagName === 'peta') targetWidth = 786;
                                        targetWidth = Math.min(targetWidth, pageMaxWidth);
                                        const ratio = targetWidth / dimensions.width;
                                        return [Math.round(targetWidth), Math.round(dimensions.height * ratio)];
                                    } catch (e) { return [150,150]; }
                                }
                            })],
                        });
                        usedImageModule = true;
                    } catch (retryErr) {
                        console.error('Retry compilation after removing tags failed:', retryErr && retryErr.properties ? retryErr.properties : retryErr);
                        renderWarnings.push('Failed to recompile template with image module after removing problematic tags. Images will be skipped.');
                        usedImageModule = false;
                        doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
                    }
                } else {
                    renderWarnings.push('Docxtemplater image-module compilation failed; images will be skipped.');
                    usedImageModule = false;
                    doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
                }
            } catch (fixErr) {
                console.error('Auto-fix for compile errors failed:', fixErr);
                renderWarnings.push('Auto-fix for template compile errors failed; images will be skipped.');
                usedImageModule = false;
                doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: () => '' });
            }
        }

    // If DEBUG enabled, log a compact summary of the image-related data we will render
        if (process.env.DEBUG) {
            try {
                const summary = {};
                Object.keys(data).forEach(k => {
                    const v = data[k];
                    if (v == null) return summary[k] = null;
                    if (Buffer.isBuffer(v)) return summary[k] = `<Buffer ${v.length} bytes>`;
                    if (typeof v === 'string') return summary[k] = v.length > 200 ? v.slice(0, 200) + '...(truncated)' : v;
                    if (Array.isArray(v)) return summary[k] = `[Array length ${v.length}]`;
                    summary[k] = Object.prototype.toString.call(v);
                });
                console.error('[DEBUG] Prepared data summary before render:', summary);
            } catch (dumpErr) {
                console.error('[DEBUG] Failed to serialize data summary:', dumpErr);
            }
        }

        // Sanitize image placeholders: docxtemplater-image-module expects tag values to be
        // strings (data URLs), Buffers, or arrays of those. If we detect an unexpected
        // object (which may be missing sizePixel and cause the module to index undefined),
        // coerce it to null. This prevents the TypeError: cannot read properties of undefined.
        try {
            const coerced = {};
            uploadFields.forEach(f => {
                const k = f.name;
                const v = data[k];
                if (v == null) return;
                // Arrays: ensure elements are string/Buffer, otherwise coerce element to null
                if (Array.isArray(v)) {
                    const newArr = v.map(el => {
                        if (el == null) return null;
                        if (Buffer.isBuffer(el)) return el;
                        if (typeof el === 'string') return el;
                        if (typeof el === 'object') {
                            // try common file-like shapes
                            if (el.buffer && Buffer.isBuffer(el.buffer)) return `data:${el.mimetype||'application/octet-stream'};base64,${el.buffer.toString('base64')}`;
                            if (el.data && Buffer.isBuffer(el.data)) return `data:${el.mimetype||'application/octet-stream'};base64,${el.data.toString('base64')}`;
                            return null;
                        }
                        return null;
                    });
                    data[k] = newArr.filter(x => x != null);
                    if (process.env.DEBUG && newArr.length !== v.length) coerced[k] = { from: v.length, to: data[k].length };
                    return;
                }
                // Single value: allow string or buffer, otherwise try to coerce file-like objects
                if (Buffer.isBuffer(v) || typeof v === 'string') return;
                if (typeof v === 'object') {
                    if (v.buffer && Buffer.isBuffer(v.buffer)) {
                        data[k] = `data:${v.mimetype||'application/octet-stream'};base64,${v.buffer.toString('base64')}`;
                        coerced[k] = 'object->dataUrl';
                        return;
                    }
                    if (v.data && Buffer.isBuffer(v.data)) {
                        data[k] = `data:${v.mimetype||'application/octet-stream'};base64,${v.data.toString('base64')}`;
                        coerced[k] = 'object->dataUrl';
                        return;
                    }
                    // Unknown object shape -> null it to avoid passing to image module
                    data[k] = null;
                    coerced[k] = 'object->null';
                }
            });
            if (process.env.DEBUG) console.error('[DEBUG] Sanitized image placeholders:', coerced);
        } catch (sanitizeErr) {
            console.error('[DEBUG] Sanitization failed:', sanitizeErr);
        }

        try {
            // Guard: Ensure any image module callbacks won't get unexpected undefined shapes
            // Wrap options.getImage/getSize to coerce undefined/falsy sizes into a safe [150,150]
            const originalGetImage = doc.options.modules && doc.options.modules[0] && doc.options.modules[0].options && doc.options.modules[0].options.getImage;
            const originalGetSize = doc.options.modules && doc.options.modules[0] && doc.options.modules[0].options && doc.options.modules[0].options.getSize;
            if (process.env.DEBUG) console.error('[DEBUG] Wrapping getImage/getSize to guard return shapes');
            if (originalGetImage || originalGetSize) {
                // replace the module's callbacks with guarded versions
                const mod = doc.options.modules[0];
                if (originalGetImage) {
                    mod.options.getImage = function(tag, name) {
                        try {
                            return originalGetImage(tag, name);
                        } catch (err) {
                            console.error('[DEBUG] getImage threw:', err);
                            return null;
                        }
                    };
                }
                if (originalGetSize) {
                    mod.options.getSize = function(imgBuffer, tag, name) {
                        try {
                            const out = originalGetSize(imgBuffer, tag, name);
                            if (!out || !Array.isArray(out) || out.length < 2 || isNaN(out[0]) || isNaN(out[1])) {
                                if (process.env.DEBUG) console.error('[DEBUG] getSize returned invalid:', out);
                                return [150, 150];
                            }
                            return out;
                        } catch (err) {
                            console.error('[DEBUG] getSize threw:', err);
                            return [150, 150];
                        }
                    };
                }
            }

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
        if (renderWarnings && renderWarnings.length > 0) {
            try { res.setHeader('X-Render-Warnings', renderWarnings.join('; ')); } catch (e) { /* ignore header set error */ }
        }
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
