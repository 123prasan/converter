const { Worker } = require('bullmq');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

// --- CONFIGURATION ---
const REDIS_CONFIG = {
    connection: {
        host: '127.0.0.1',
        port: 6378 // Matching port
    }
};

const uploadDir = path.resolve(__dirname, 'uploads');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);

// ============================================================================
//  PYTHON EXECUTOR
// ============================================================================
async function runPython(scriptName, args, job, socketId) {
    return new Promise((resolve, reject) => {
        console.log(`[EXEC] python ${scriptName} ${args.map(a => path.basename(a)).join(' ')}`);

        const python = spawn('python', [scriptName, ...args], {
            cwd: __dirname,
            env: { ...process.env, PYTHONUTF8: '1' }
        });

        let fullOutput = "";
        let stderrOutput = "";

        python.stdout.on('data', (data) => {
            const msg = data.toString();
            fullOutput += msg;
            const trimmed = msg.trim();
            if (trimmed) job.updateProgress({ message: trimmed, socketId });
        });

        python.stderr.on('data', (data) => { stderrOutput += data.toString(); });

        python.on('close', (code) => {
            // Delete inputs after processing to keep disk clean
            // Note: In production you might want to wait a bit or let TTL handle it
            
            if (code === 0) {
                const timeMatch = fullOutput.match(/won in ([\d\.]+)s/);
                resolve({ duration: timeMatch ? timeMatch[1] : "0.00" });
            } else {
                reject(new Error(`Exit code ${code}: ${stderrOutput.slice(0, 200)}`));
            }
        });
    });
}

// ============================================================================
//  THE WORKER LOGIC
// ============================================================================
console.log("👷 Worker System Online (Port 6378). Waiting for jobs...");

const worker = new Worker('conversion-queue', async (job) => {
    const { inputPath, socketId } = job.data;
    const jobName = job.name;
    
    console.log(`[JOB START] ${jobName} (ID: ${job.id})`);

    let script = "";
    let args = [];
    let outputFilename = "";

    // --- JOB ROUTING TABLE ---
    switch (jobName) {
        case 'word-to-pdf':
            script = 'word_to_pdf.py';
            outputFilename = `word_${Date.now()}.pdf`;
            args = [inputPath, path.join(uploadDir, outputFilename)];
            break;

        case 'pdf-to-doc':
            script = 'convert.py';
            outputFilename = `converted_${Date.now()}.docx`;
            args = [inputPath, job.data.tool, path.join(uploadDir, outputFilename), job.data.secretKey];
            break;

        case 'ppt-to-pdf':
            script = 'ppt_to_pdf.py';
            outputFilename = `ppt_${Date.now()}.pdf`;
            args = [inputPath, path.join(uploadDir, outputFilename)];
            break;

        case 'image-to-pdf':
            script = 'image_to_pdf.py';
            outputFilename = `image_${Date.now()}.pdf`;
            args = [inputPath, path.join(uploadDir, outputFilename)];
            break;

        case 'ocr-extract':
            script = 'ocr_engine.py';
            outputFilename = `ocr_${Date.now()}.txt`;
            args = [inputPath, path.join(uploadDir, outputFilename)];
            break;

        case 'compress-pdf':
            script = 'compressor.py';
            outputFilename = `compressed_${Date.now()}.pdf`;
            args = [inputPath, path.join(uploadDir, outputFilename), job.data.grayscale, 'true', job.data.dpi];
            break;

        case 'compress-docx':
            script = 'compress_docx.py';
            outputFilename = `compressed_${Date.now()}.docx`;
            args = [inputPath, path.join(uploadDir, outputFilename), job.data.level];
            break;

        case 'compress-image':
            script = 'compress_image.py';
            outputFilename = `compressed_${Date.now()}.${job.data.ext}`;
            args = [inputPath, path.join(uploadDir, outputFilename), job.data.level, job.data.ext];
            break;

        case 'sign-pdf':
            script = 'sign_pdf.py';
            outputFilename = `signed_${Date.now()}.pdf`;
            const p = job.data.params;
            args = [
                inputPath, path.join(uploadDir, outputFilename),
                p.name, p.pageRange, p.x, p.y, p.fontSize, p.textColor, p.fontStyle, p.opacity
            ];
            break;

        case 'merge-pdf':
            script = 'merge_pdf.py';
            outputFilename = `merged_${Date.now()}.pdf`;
            args = [path.join(uploadDir, outputFilename), ...job.data.inputPaths];
            break;
    }

    if (!script) throw new Error(`Unknown job type: ${jobName}`);

    // --- EXECUTE ---
    try {
        const result = await runPython(script, args, job, socketId);
        console.log(`[JOB DONE] ${jobName} finished.`);
        
        return {
            socketId: socketId,
            result: {
                status: 'success',
                downloadUrl: `/download/${outputFilename}`,
                duration: result.duration,
                file: outputFilename
            }
        };
    } catch (err) {
        job.updateProgress({ message: "Error: " + err.message, socketId });
        throw err;
    }

}, { 
    connection: REDIS_CONFIG.connection, 
    concurrency: 5 
});