const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const http = require('http');
const { Server } = require('socket.io');
const { Queue, QueueEvents } = require('bullmq');
const net = require('net');
const { spawn, spawnSync } = require('child_process');
const compression = require('compression');

// --- CONFIGURATION ---
const REDIS_CONFIG = {
    connection: {
        host: '127.0.0.1',
        port: 6378 // Matches your specific Redis port
    }
};

const app = express();
const server = http.createServer(app);

// Enable CORS
app.use(cors({ origin: '*', methods: ['GET', 'POST'] }));

const io = new Server(server, {
    cors: { origin: "*", methods: ["GET", "POST"] }
});

// --- QUEUE SETUP ---
const conversionQueue = new Queue('conversion-queue', REDIS_CONFIG);
const queueEvents = new QueueEvents('conversion-queue', REDIS_CONFIG);

// --- GLOBAL EVENT LISTENERS (Bridge Worker -> Frontend) ---
queueEvents.on('completed', ({ jobId, returnvalue }) => {
    if (returnvalue && returnvalue.socketId) {
        io.to(returnvalue.socketId).emit('conversion-finished', returnvalue.result);
    }
});

queueEvents.on('progress', ({ jobId, data }) => {
    if (data && data.socketId) {
        io.to(data.socketId).emit('conversion-status', data.message);
    }
});

queueEvents.on('failed', ({ jobId, failedReason }) => {
    console.error(`Job ${jobId} failed: ${failedReason}`);
});

// --- DIRECTORIES & CLEANUP ---
const uploadDir = path.resolve(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

const outputsDir = path.resolve(__dirname, 'outputs');
if (!fs.existsSync(outputsDir)) fs.mkdirSync(outputsDir);

const FILE_TTL_MS = parseInt(process.env.FILE_TTL_MS, 10) || 5 * 60 * 1000; // 5 mins
const CLEANUP_INTERVAL_MS = parseInt(process.env.CLEANUP_INTERVAL_MS, 10) || 60 * 1000; // 1 min
const deletionTimers = new Map();

// Helper to schedule deletion (Files live for TTL)
const scheduleDeletion = (filePath, ttl = FILE_TTL_MS) => {
    if (!filePath || deletionTimers.has(filePath)) return;
    try { if (!fs.existsSync(filePath)) return; } catch (e) { return; }

    const timer = setTimeout(() => {
        fs.unlink(filePath, (err) => {
            if (!err) console.log(`[CLEANUP] Deleted ${path.basename(filePath)}`);
            deletionTimers.delete(filePath);
        });
    }, ttl);
    deletionTimers.set(filePath, timer);
};

// Periodic Cleanup Scanner
const scanAndCleanup = async (dir) => {
    try {
        const files = await fs.promises.readdir(dir);
        const now = Date.now();
        await Promise.all(files.map(async (file) => {
            const filePath = path.join(dir, file);
            try {
                const stats = await fs.promises.stat(filePath);
                if (stats.isFile() && (now - Math.max(stats.mtimeMs, stats.ctimeMs) > FILE_TTL_MS)) {
                    if (!deletionTimers.has(filePath)) fs.unlink(filePath, () => {});
                }
            } catch (e) {}
        }));
    } catch (e) {}
};

setInterval(() => { scanAndCleanup(uploadDir); scanAndCleanup(outputsDir); }, CLEANUP_INTERVAL_MS);
scanAndCleanup(uploadDir);
scanAndCleanup(outputsDir);

// --- MIDDLEWARE ---
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(compression());
app.use(express.static(path.join(__dirname, 'public'), { maxAge: '30d' }));
app.use('/uploads', express.static(uploadDir));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// ============================================================================
//  SEO MIDDLEWARE (The Fix is Here)
// ============================================================================
app.use((req, res, next) => {
    const siteName = 'OMNIGrid NX';
    const baseUrl = `${req.protocol}://${req.get('host')}`;

    // Initialize the objects your EJS template expects
    res.locals.meta = res.locals.meta || {};
    res.locals.meta.title = res.locals.meta.title || `${siteName} — Secure Online Converters & PDF Tools`;
    res.locals.meta.description = res.locals.meta.description || 'Fast, secure online PDF & document converters. Convert, compress, merge, sign, and more — privacy-first, files auto-deleted after 5 minutes.';
    res.locals.meta.url = baseUrl + req.originalUrl;
    res.locals.meta.image = res.locals.meta.image || `${baseUrl}/assets/og-image.svg`;
    res.locals.meta.type = res.locals.meta.type || 'website';
    res.locals.meta.keywords = res.locals.meta.keywords || 'pdf converter, word to pdf, ppt to pdf, compress pdf, merge pdf, online converters';

    res.locals.canonicalUrl = (res.locals.meta.url || baseUrl).split('?')[0];
    
    // Ensure jsonLd exists (this fixes the error)
    res.locals.jsonLd = res.locals.jsonLd || null;
    
    // Ensure extraHead exists
    res.locals.extraHead = res.locals.extraHead || '';
    
    res.locals.gaId = process.env.GA_TRACKING_ID || null;

    next();
});

// Per-route SEO overrides (Add specific titles for tools)
app.use((req, res, next) => {
    const map = {
        '/word-to-pdf': { title: 'Word to PDF - Free Online Converter', description: 'Convert Word (.doc, .docx) to high-quality PDF instantly. Secure, fast, and auto-deleted uploads.' },
        '/image-to-pdf': { title: 'Image to PDF - Free Converter', description: 'Combine images into a single PDF quickly. Supports PNG, JPG and more.' },
        '/compress': { title: 'Compress PDF - Reduce File Size', description: 'Reduce PDF size while retaining quality. Fast online compression with privacy-first policy.' },
        '/merge': { title: 'Merge PDF - Combine Documents Online', description: 'Merge multiple PDFs into one document quickly and securely.' },
        '/sign-upload': { title: 'Sign PDF - Online Signature & Watermark', description: 'Add signatures and watermarks to PDFs with precise positioning and style controls.' }
    };

    if (map[req.path]) {
        res.locals.meta = Object.assign({}, res.locals.meta, map[req.path]);
        res.locals.canonicalUrl = `${req.protocol}://${req.get('host')}${req.path}`;
        
        // Attach JSON-LD snippet for tool pages
        res.locals.jsonLd = {
            "@context": "https://schema.org",
            "@type": "SoftwareApplication",
            "name": res.locals.meta.title,
            "operatingSystem": "Web",
            "applicationCategory": "Utilities",
            "description": res.locals.meta.description,
            "url": res.locals.canonicalUrl
        };
    }
    next();
});

// --- UPLOAD CONFIG ---
const upload = multer({ 
    storage: multer.diskStorage({
        destination: (req, file, cb) => cb(null, uploadDir),
        filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
    }),
    limits: { fileSize: 50 * 1024 * 1024 } // 50MB
});

// ============================================================================
//  QUEUED ROUTES (All Features Ported)
// ============================================================================

// 1. Word to PDF
app.post('/word-to-pdf-convert', upload.single('word'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('word-to-pdf', {
        inputPath: req.file.path,
        socketId: req.body.socketId
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// 2. PPT to PDF
app.post('/ppt-to-pdf-convert', upload.single('ppt'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('ppt-to-pdf', {
        inputPath: req.file.path,
        socketId: req.body.socketId
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// 3. PDF to Doc (E2EE)
app.post('/convert', upload.single('pdf'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('pdf-to-doc', {
        inputPath: req.file.path,
        tool: (req.body.tool || 'pdf2doc').toLowerCase(),
        secretKey: req.body.secretKey || '',
        socketId: req.body.socketId
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// 4. Merge PDF (Array)
app.post('/merge-pdf', upload.array('pdfs', 20), async (req, res) => {
    if (!req.files || req.files.length < 2) return res.status(400).json({ error: "Need 2+ files" });
    
    req.files.forEach(f => scheduleDeletion(f.path));
    
    await conversionQueue.add('merge-pdf', {
        inputPaths: req.files.map(f => f.path),
        socketId: req.body.socketId
    });
    res.json({ status: "queued", message: "Queued..." });
});

// 5. Sign PDF (Complex Args)
app.post('/sign-pdf', upload.fields([{ name: 'pdf', maxCount: 1 }]), async (req, res) => {
    const file = req.files['pdf'] ? req.files['pdf'][0] : null;
    if (!file) return res.status(400).json({ error: "No PDF" });
    
    scheduleDeletion(file.path);

    await conversionQueue.add('sign-pdf', {
        inputPath: file.path,
        socketId: req.body.socketId,
        params: {
            name: req.body.name || "Sign",
            pageRange: req.body.pageRange || "all",
            x: req.body.x || "50",
            y: req.body.y || "50",
            fontSize: req.body.fontSize || "24",
            textColor: req.body.textColor || "#000",
            fontStyle: req.body.fontStyle || "script",
            opacity: req.body.opacity || "1"
        }
    });
    res.json({ status: "queued", message: "Queued..." });
});

// 6. Image to PDF
app.post('/image-to-pdf-convert', upload.single('image'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('image-to-pdf', {
        inputPath: req.file.path,
        socketId: req.body.socketId
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// 7. OCR
app.post('/ocr-extract', upload.single('image'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No image" });
    await conversionQueue.add('ocr-extract', {
        inputPath: req.file.path,
        socketId: req.body.socketId
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// 8. Compress PDF
app.post('/compress-pdf', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('compress-pdf', {
        inputPath: req.file.path,
        socketId: req.body.socketId,
        grayscale: req.body.grayscale || 'false',
        dpi: req.body.dpi || '150'
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// 9. Compress Docx
app.post('/compress-docx', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('compress-docx', {
        inputPath: req.file.path,
        socketId: req.body.socketId,
        level: req.body.level || 'recommended'
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// 10. Compress Image
app.post('/compress-image', upload.single('file'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('compress-image', {
        inputPath: req.file.path,
        socketId: req.body.socketId,
        level: req.body.level || 'recommended',
        ext: path.extname(req.file.originalname).toLowerCase().slice(1)
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// Legacy route alias
app.post('/compress', upload.single('pdf'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    await conversionQueue.add('compress-pdf', {
        inputPath: req.file.path,
        socketId: req.body.socketId,
        grayscale: req.body.grayscale || 'false',
        dpi: req.body.dpi || '150'
    });
    scheduleDeletion(req.file.path);
    res.json({ status: "queued", message: "Queued..." });
});

// --- HELPER ROUTES ---
app.get('/download/:filename', (req, res) => {
    const file = path.join(uploadDir, req.params.filename);
    res.download(file, (err) => { if (err) res.status(404).send("Expired"); });
});

app.get('/', (req, res) => res.render('index'));

const views = ['ptod', 'image-pdf', 'compress', 'ocrtext', 'unlock', 'exceltopdf', 'protectpdf', 'mergepdf', 'singn', 'pdf2jpg', 'ppttopdf', 'doc-pdf'];
const routeAliases = { 'ptod': '/pdf-to-doc', 'doc-pdf': '/word-to-pdf', 'image-pdf': '/image-to-pdf', 'compress': '/compress', 'ocrtext': '/ocr', 'exceltopdf': '/excel-to-pdf', 'protectpdf': '/protect', 'mergepdf': '/merge', 'singn': '/sign', 'pdf2jpg': '/pdf-to-jpg', 'ppttopdf': '/ppt-to-pdf', 'unlock': '/unlock' };

views.forEach(v => {
    app.get(`/${v}`, (req, res) => res.render(v));
    if (routeAliases[v]) app.get(routeAliases[v], (req, res) => res.render(v));
});

// Editor Routes
app.get('/sign-upload', (req, res) => res.render('sign-upload'));
app.post('/upload-pdf-for-editor', upload.single('pdf'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: 'No file' });
    scheduleDeletion(req.file.path);
    res.json({ path: `/uploads/${req.file.filename}`, filename: req.file.originalname });
});
app.get('/sign-watermark-editor', (req, res) => {
    const pdfQuery = req.query.pdf;
    if (!pdfQuery) return res.redirect('/sign-upload');
    res.render('sign-watermark-editor', { pdfPath: pdfQuery, filename: req.query.name || 'doc.pdf' });
});

app.get('/cleanup-status', async (req, res) => res.json({ status: "Active via Queue", scheduled: deletionTimers.size }));

app.get('/robots.txt', (req, res) => {
    const base = `${req.protocol}://${req.get('host')}`;
    res.type('text/plain').send(`User-agent: *\nAllow: /\nSitemap: ${base}/sitemap.xml\n`);
});

app.get('/sitemap.xml', (req, res) => {
    const base = `${req.protocol}://${req.get('host')}`;
    const pages = ['/', '/word-to-pdf', '/image-to-pdf', '/compress', '/merge', '/sign-upload', '/pdf-to-doc', '/excel-to-pdf', '/ppt-to-pdf', '/protect', '/ocr', '/pdf-to-jpg'];
    const urls = pages.map(p => `  <url>\n    <loc>${base + p}</loc>\n    <lastmod>${new Date().toISOString()}</lastmod>\n  </url>`).join('\n');
    res.type('application/xml').send(`<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n${urls}\n</urlset>`);
});

// Global Error Handler
app.use((err, req, res, next) => {
    if (err instanceof multer.MulterError) return res.status(400).json({ error: err.message });
    console.error(err);
    if (!res.headersSent) res.status(500).send("Server Error");
});

// --- SERVER START ---
const PORT = process.env.PORT || 3000;

function startLibreOfficeDaemon() {
    const bin = process.env.LIBREOFFICE_BIN || 'soffice';
    const client = new net.Socket();
    client.setTimeout(1000);
    client.on('error', () => {
        try {
            const check = spawnSync(process.platform === 'win32' ? 'where' : 'which', [bin]);
            if (check.status !== 0) return console.warn(`LibreOffice not found: ${bin}`);
            
            const child = spawn(bin, ['--headless', '--norestore', '--invisible', '--accept=socket,host=127.0.0.1,port=2002;urp;'], {
                detached: true, stdio: 'ignore'
            });
            child.unref();
            console.log("Started LibreOffice Daemon on 2002");
        } catch (e) { console.error("Failed to start LibreOffice:", e.message); }
    });
    client.connect(2002, '127.0.0.1', () => client.end());
}

startLibreOfficeDaemon();

server.listen(PORT, () => {
    console.log(`🚀 API Gateway (Producer) running on port ${PORT}`);
});