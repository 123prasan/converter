const express = require('express');
const multer = require('multer');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const http = require('http');
const { Server } = require('socket.io');
const net = require('net');

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
    cors: { origin: "*", methods: ["GET", "POST"] }
});

const uploadDir = path.resolve(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

// Ensure outputs directory (some scripts write to outputs/ or we want to clean it too)
const outputsDir = path.resolve(__dirname, 'outputs');
if (!fs.existsSync(outputsDir)) fs.mkdirSync(outputsDir);

// File TTL and cleanup interval (configurable via env vars). Default TTL = 5 minutes, scan every 1 minute.
const FILE_TTL_MS = parseInt(process.env.FILE_TTL_MS, 10) || 5 * 60 * 1000; // 5 minutes
const CLEANUP_INTERVAL_MS = parseInt(process.env.CLEANUP_INTERVAL_MS, 10) || 60 * 1000; // 1 minute

// Map of scheduled deletion timers so we can avoid double-scheduling and inspect status
const deletionTimers = new Map();

const scheduleDeletion = (filePath, ttl = FILE_TTL_MS) => {
    if (!filePath) return;
    if (deletionTimers.has(filePath)) return; // already scheduled

    try {
        if (!fs.existsSync(filePath)) return; // nothing to delete
    } catch (e) {
        return;
    }

    const timer = setTimeout(() => {
        fs.unlink(filePath, (err) => {
            if (err) {
                console.error(`[CLEANUP] Failed to delete ${filePath}:`, err.message);
            } else {
                console.log(`[CLEANUP] Deleted ${filePath}`);
            }
            deletionTimers.delete(filePath);
        });
    }, ttl);

    deletionTimers.set(filePath, timer);
    console.log(`[CLEANUP] Scheduled ${filePath} for deletion in ${Math.round(ttl/1000)}s`);
};

// Utility: scan a directory and remove files older than TTL
const scanAndCleanup = async (dir) => {
    try {
        const files = await fs.promises.readdir(dir);
        const now = Date.now();
        await Promise.all(files.map(async (file) => {
            const filePath = path.join(dir, file);
            let stats;
            try { stats = await fs.promises.stat(filePath); } catch (e) { return; }
            if (!stats.isFile()) return;
            const age = now - Math.max(stats.mtimeMs || 0, stats.ctimeMs || 0);
            if (age > FILE_TTL_MS) {
                // If not already scheduled, delete immediately
                if (!deletionTimers.has(filePath)) {
                    try {
                        await fs.promises.unlink(filePath);
                        console.log(`[CLEANUP] Purged ${filePath} (age ${Math.round(age/1000)}s)`);
                    } catch (e) {
                        console.error(`[CLEANUP] Failed purge of ${filePath}:`, e.message);
                    }
                }
            }
        }));
    } catch (e) {
        // ignore
    }
};

// Periodic scanner to catch files that may have been missed (e.g., server restarted or timers cleared)
setInterval(() => {
    scanAndCleanup(uploadDir);
    scanAndCleanup(outputsDir);
}, CLEANUP_INTERVAL_MS);

// Initial scan on startup to remove anything older than TTL
scanAndCleanup(uploadDir);
scanAndCleanup(outputsDir);

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
// Serve uploaded files over HTTP so PDF.js can fetch them (do NOT use file:/// paths)
app.use('/uploads', express.static(uploadDir));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// --- STORAGE CONFIGURATION ---
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadDir),
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, uniqueSuffix + '-' + file.originalname);
    }
});
const upload = multer({ storage, limits: { fileSize: 50 * 1024 * 1024 } }); // 50MB limit

// --- UNIFIED EXECUTION ENGINE ---
/**
 * Generic function to run Python scripts and handle Socket + HTTP responses
 */
function executeTask(res, options) {
    const { script, args, socketId, inputPaths, outputFilename } = options;
    const outputPath = path.join(uploadDir, outputFilename);

    console.log(`[EXEC] python ${script} ${args.join(' ')}`);

    const python = spawn('python', [script, ...args], {
        cwd: __dirname,
        env: { ...process.env, PYTHONUTF8: '1' }
    });

    python.on('error', (err) => {
        console.error(`[SPAWN ERROR] ${script}:`, err.message);
        if (socketId) io.to(socketId).emit('conversion-error', 'Engine spawn failed: ' + err.message);
        if (!res.headersSent) res.status(500).json({ error: 'Engine spawn failed', details: err.message });
    });

    python.stdout.on('data', (data) => {
        const msg = data.toString().trim();
        if (msg) {
            console.log(`[PYTHON] ${script}: ${msg}`);
            if (socketId) io.to(socketId).emit('conversion-status', msg);
        }
    });

    // Capture stderr for richer diagnostics
    let stderrBuf = '';
    python.stderr.on('data', (data) => {
        const err = data.toString();
        stderrBuf += err;
        const trimmed = err.trim();
        if (trimmed) console.error(`[PYTHON ERROR] ${script}:`, trimmed);
    });

    python.on('close', (code) => {
        // 1. Schedule input files for deletion (instead of immediate unlink). This ensures files live for FILE_TTL_MS and makes behavior robust across restarts.
        inputPaths.forEach(p => scheduleDeletion(p));

        if (code === 0) {
            const stats = fs.existsSync(outputPath) ? fs.statSync(outputPath) : { size: 0 };
            const responseData = { 
                status: 'success',
                downloadUrl: `/download/${outputFilename}`,
                fileSize: stats.size 
            };
            
            // Send final socket update
            if (socketId) io.to(socketId).emit('conversion-finished', responseData);

            // Schedule output file for deletion after TTL
            scheduleDeletion(outputPath);
            
            // If the response hasn't been sent yet (sync request), send it
            if (!res.headersSent) res.json(responseData);
        } else {
            console.error(`[PROCESS EXIT] ${script} failed with code ${code}`);
            const details = stderrBuf || 'No stderr output captured.';
            // If output was created despite failure, schedule it as well
            if (fs.existsSync(outputPath)) scheduleDeletion(outputPath);
            
            if (socketId) io.to(socketId).emit('conversion-error', `Engine failed to process. Details: ${details.slice(0,1000)}`);
            if (!res.headersSent) res.status(500).json({ error: "Processing failed", code, details: details.slice(0,2000) });
        }
    });
}

// --- ROUTES ---
app.get('/', (req, res) => res.render('index'));
const views = ['ptod', 'image-pdf', 'compress', 'ocrtext', 'unlock', 'exceltopdf', 'protectpdf', 'mergepdf', 'singn', 'pdf2jpg', 'ppttopdf', 'doc-pdf'];
const routeAliases = {
  'ptod': '/pdf-to-doc',
  'doc-pdf': '/word-to-pdf',
  'image-pdf': '/image-to-pdf',
  'compress': '/compress',
  'ocrtext': '/ocr',
  'exceltopdf': '/excel-to-pdf',
  'protectpdf': '/protect',
  'mergepdf': '/merge',
  'singn': '/sign',
  'pdf2jpg': '/pdf-to-jpg',
  'ppttopdf': '/ppt-to-pdf',
  'unlock': '/unlock'
};

views.forEach(v => {
  app.get(`/${v}`, (req, res) => res.render(v));
  if (routeAliases[v]) app.get(routeAliases[v], (req, res) => res.render(v));
});

// --- DOWNLOAD UTILITY ---
app.get('/download/:filename', (req, res) => {
    const file = path.join(uploadDir, req.params.filename);
    res.download(file, (err) => {
        if (err && !res.headersSent) res.status(404).send("File expired.");
    });
});

// --- TOOL ENDPOINTS ---

app.post('/word-to-pdf-convert', upload.single('word'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    const outputFilename = `word_${Date.now()}.pdf`;
    executeTask(res, {
        script: 'word_to_pdf.py',
        // Pass raw filesystem paths (no extra quoting) so Python receives clean argv values
        args: [path.resolve(req.file.path), path.join(uploadDir, outputFilename)],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

app.post('/ppt-to-pdf-convert', upload.single('ppt'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    const outputFilename = `ppt_${Date.now()}.pdf`;
    executeTask(res, {
        script: 'ppt_to_pdf.py',
        args: [path.resolve(req.file.path), path.join(uploadDir, outputFilename)],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

app.post('/sign-pdf', upload.fields([{ name: 'pdf', maxCount: 1 }]), (req, res) => {
    const file = req.files['pdf'] ? req.files['pdf'][0] : null;
    if (!file) return res.status(400).json({ error: "No PDF" });
    const outputFilename = `signed_${Date.now()}.pdf`;
    const { name, pageRange, fontSize, textColor, fontStyle, socketId, x, y, opacity } = req.body;

    executeTask(res, {
        script: 'sign_pdf.py',
        args: [
            path.resolve(file.path), 
            path.join(uploadDir, outputFilename),
            name || "Sign", 
            pageRange || "all", 
            x || "50", 
            y || "50", 
            fontSize || "24", 
            textColor || "#000", 
            fontStyle || "script",
            opacity || "1"
        ],
        socketId,
        inputPaths: [file.path],
        outputFilename
    });
});

app.post('/merge-pdf', upload.array('pdfs', 20), (req, res) => {
    if (!req.files || req.files.length < 2) return res.status(400).json({ error: "Need 2+ files" });
    const outputFilename = `merged_${Date.now()}.pdf`;
    executeTask(res, {
        script: 'merge_pdf.py',
        args: [path.join(uploadDir, outputFilename), ...req.files.map(f => path.resolve(f.path))],
        socketId: req.body.socketId,
        inputPaths: req.files.map(f => f.path),
        outputFilename
    });
});

// PDF to Doc (E2EE-aware) - CRITICAL ROUTE FOR PTOD PAGE
app.post('/convert', upload.single('pdf'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    
    const tool = (req.body.tool || 'pdf2doc').toLowerCase();
    const secretKey = req.body.secretKey || '';
    const outputFilename = `converted_${Date.now()}.docx`;

    // Validate secret key for E2EE
    if (!secretKey) {
        return res.status(400).json({ error: "Missing encryption key" });
    }

    executeTask(res, {
        script: 'convert.py',
        args: [
            path.resolve(req.file.path),
            tool,
            path.join(uploadDir, outputFilename),
            secretKey
        ],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

// Image to PDF
app.post('/image-to-pdf-convert', upload.single('image'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file" });
    const outputFilename = `image_${Date.now()}.pdf`;

    executeTask(res, {
        script: 'image_to_pdf.py',
        args: [path.resolve(req.file.path), path.join(uploadDir, outputFilename)],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

// OCR extract (image -> txt)
app.post('/ocr-extract', upload.single('image'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No image provided" });
    const outputFilename = `ocr_${Date.now()}.txt`;

    executeTask(res, {
        script: 'ocr_engine.py',
        args: [path.resolve(req.file.path), path.join(uploadDir, outputFilename)],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

// PDF Compression
app.post('/compress-pdf', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const outputFilename = `compressed_${Date.now()}.pdf`;
    const grayscale = req.body.grayscale || 'false';
    const dpi = req.body.dpi || '150';

    executeTask(res, {
        script: 'compressor.py',
        args: [
            path.resolve(req.file.path),
            path.join(uploadDir, outputFilename),
            grayscale,
            'true',  // strip_meta
            dpi
        ],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

// DOCX Compression
app.post('/compress-docx', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const outputFilename = `compressed_${Date.now()}.docx`;
    const level = req.body.level || 'recommended';

    executeTask(res, {
        script: 'compress_docx.py',
        args: [
            path.resolve(req.file.path),
            path.join(uploadDir, outputFilename),
            level
        ],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

// Image Compression
app.post('/compress-image', upload.single('file'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const ext = path.extname(req.file.originalname).toLowerCase().slice(1);
    const outputFilename = `compressed_${Date.now()}.${ext}`;
    const level = req.body.level || 'recommended';

    executeTask(res, {
        script: 'compress_image.py',
        args: [
            path.resolve(req.file.path),
            path.join(uploadDir, outputFilename),
            level,
            ext
        ],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

// Legacy /compress endpoint (redirect to PDF for backward compatibility)
app.post('/compress', upload.single('pdf'), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const outputFilename = `compressed_${Date.now()}.pdf`;
    const grayscale = req.body.grayscale || 'false';
    const dpi = req.body.dpi || '150';

    executeTask(res, {
        script: 'compressor.py',
        args: [
            path.resolve(req.file.path),
            path.join(uploadDir, outputFilename),
            grayscale,
            'true',  // strip_meta
            dpi
        ],
        socketId: req.body.socketId,
        inputPaths: [req.file.path],
        outputFilename
    });
});

// --- SIGN/WATERMARK EDITOR ROUTES ---
app.get('/sign-upload', (req, res) => {
    res.render('sign-upload');
});

app.post('/upload-pdf-for-editor', upload.single('pdf'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
    }
    // Return a web-accessible path under /uploads so the client can load it over HTTP
    const webPath = `/uploads/${req.file.filename}`;

    // Schedule uploaded file for deletion within configured TTL
    scheduleDeletion(req.file.path);

    res.json({ path: webPath, filename: req.file.originalname });
});

app.get('/sign-watermark-editor', (req, res) => {
    const pdfQuery = req.query.pdf;
    const filename = req.query.name || 'document.pdf';

    if (!pdfQuery) {
        return res.redirect('/sign-upload');
    }

    // Normalize possible inputs: web path (/uploads/...), absolute disk path, or bare filename
    let pdfPathOnDisk;
    let webPath;

    if (pdfQuery.startsWith('/uploads/')) {
        const fileName = path.basename(pdfQuery);
        pdfPathOnDisk = path.join(uploadDir, fileName);
        webPath = pdfQuery;
    } else if (path.isAbsolute(pdfQuery)) {
        pdfPathOnDisk = pdfQuery;
        webPath = `/uploads/${path.basename(pdfQuery)}`;
    } else {
        pdfPathOnDisk = path.join(uploadDir, pdfQuery);
        webPath = `/uploads/${path.basename(pdfQuery)}`;
    }

    if (!fs.existsSync(pdfPathOnDisk)) {
        return res.redirect('/sign-upload');
    }

    res.render('sign-watermark-editor', { pdfPath: webPath, filename });
});

// --- CLEANUP STATUS (debugging) ---
app.get('/cleanup-status', async (req, res) => {
    const gather = async (dir) => {
        try {
            const files = await fs.promises.readdir(dir);
            return await Promise.all(files.map(async (f) => {
                const p = path.join(dir, f);
                try {
                    const s = await fs.promises.stat(p);
                    return { path: p, size: s.size, mtime: s.mtimeMs, scheduled: deletionTimers.has(p) };
                } catch (e) { return null; }
            }));
        } catch (e) { return [] }
    };

    const uploads = (await gather(uploadDir)).filter(Boolean);
    const outputs = (await gather(outputsDir)).filter(Boolean);
    res.json({ uploads, outputs, ttlMs: FILE_TTL_MS, scanIntervalMs: CLEANUP_INTERVAL_MS, scheduledCount: deletionTimers.size });
});

// --- SERVER INIT ---
const PORT = process.env.PORT || 3000;

// Start a persistent LibreOffice daemon to reduce cold-start for conversions
function startLibreOfficeDaemon() {
    const host = '127.0.0.1';
    const port = process.env.LIBREOFFICE_PORT || '2002';
    const bin = process.env.LIBREOFFICE_BIN || 'soffice';

    // Quick check: if a listener is already present, do nothing
    const client = new net.Socket();
    client.setTimeout(1000);

    client.on('error', () => {
        // Not running - check if binary is available first
        const checkCmd = process.platform === 'win32' ? 'where' : 'which';
        try {
            const check = spawnSync(checkCmd, [bin], { encoding: 'utf8' });
            if (check.status !== 0) {
                console.warn(`LibreOffice binary not found in PATH: ${bin} (checked with ${checkCmd})`);
                return;
            }
        } catch (e) {
            console.warn('Error while checking for LibreOffice binary:', e.message);
            return;
        }

        // Attempt to spawn a detached soffice daemon but handle async spawn errors
        try {
            const child = spawn(bin, ['--headless', '--norestore', '--invisible', `--accept=socket,host=${host},port=${port};urp;StarOffice.ServiceManager`], {
                detached: true,
                stdio: ['ignore', 'ignore', 'pipe']
            });

            child.on('error', (err) => {
                console.error('LibreOffice daemon spawn error:', err && err.message ? err.message : err);
            });
            child.stderr.on('data', (d) => console.warn('LibreOffice stderr:', d.toString().trim()));
            child.unref();
            console.log(`Started LibreOffice daemon (${bin}) on port ${port}`);
        } catch (e) {
            console.error('Failed to start LibreOffice daemon (sync):', e.message);
        }
    });

    client.connect({ port: parseInt(port, 10), host }, () => {
        console.log(`LibreOffice daemon already listening on ${host}:${port}`);
        client.end();
    });
    client.on('timeout', () => client.destroy());
}

startLibreOfficeDaemon();

server.listen(PORT, () => console.log(`🚀 OMNIGrid NX Active on port ${PORT}`));

io.on('connection', (socket) => console.log(`Client Connected: ${socket.id}`));