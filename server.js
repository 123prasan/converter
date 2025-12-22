const express = require('express');
const multer = require('multer');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const http = require('http');
const { Server } = require('socket.io');

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
    cors: { origin: "*", methods: ["GET", "POST"] }
});

const uploadDir = path.resolve(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
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

    const python = spawn('python', [script, ...args], {
        shell: true, // Required for Windows path compatibility
        env: { ...process.env, PYTHONUTF8: '1' }
    });

    python.stdout.on('data', (data) => {
        const msg = data.toString().trim();
        if (socketId) io.to(socketId).emit('conversion-status', msg);
    });

    python.stderr.on('data', (data) => {
        console.error(`[PYTHON ERROR] ${script}:`, data.toString());
    });

    python.on('close', (code) => {
        // 1. Cleanup input files
        inputPaths.forEach(p => fs.unlink(p, () => {}));

        if (code === 0) {
            const stats = fs.existsSync(outputPath) ? fs.statSync(outputPath) : { size: 0 };
            const responseData = { 
                status: 'success',
                downloadUrl: `/download/${outputFilename}`,
                fileSize: stats.size 
            };
            
            // Send final socket update
            if (socketId) io.to(socketId).emit('conversion-finished', responseData);
            
            // If the response hasn't been sent yet (sync request), send it
            if (!res.headersSent) res.json(responseData);
        } else {
            if (socketId) io.to(socketId).emit('conversion-error', 'Engine failed to process.');
            if (!res.headersSent) res.status(500).json({ error: "Processing failed" });
        }
    });
}

// --- ROUTES ---
app.get('/', (req, res) => res.render('index'));
const views = ['ptod', 'image-pdf', 'compress', 'ocrtext', 'unlock', 'exceltopdf', 'protectpdf', 'mergepdf', 'singn', 'pdf2jpg', 'ppttopdf', 'doc-pdf'];
views.forEach(v => app.get(`/${v}`, (req, res) => res.render(v)));

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
        args: [`"${path.resolve(req.file.path)}"`, `"${path.join(uploadDir, outputFilename)}"`],
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
    const { name, pageRange, fontSize, textColor, fontStyle, socketId } = req.body;

    executeTask(res, {
        script: 'sign_pdf.py',
        args: [
            path.resolve(file.path), 
            path.join(uploadDir, outputFilename),
            name || "Sign", pageRange || "all", "50", "50", fontSize || "24", textColor || "#000", fontStyle || "bold"
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

// --- AUTO-CLEANUP (Every 30 mins) ---
setInterval(() => {
    fs.readdir(uploadDir, (err, files) => {
        if (err) return;
        const now = Date.now();
        files.forEach(file => {
            const filePath = path.join(uploadDir, file);
            fs.stat(filePath, (err, stats) => {
                if (!err && (now - new Date(stats.ctime).getTime() > 3600000)) {
                    fs.unlink(filePath, () => {});
                }
            });
        });
    });
}, 1800000);

// --- SERVER INIT ---
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => console.log(`🚀 OMNIGrid NX Active on port ${PORT}`));

io.on('connection', (socket) => console.log(`Client Connected: ${socket.id}`));