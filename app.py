"""
FastAPI web app for the MTG Packing Slip Organizer.
Accepts PDF uploads and returns organized HTML pull sheets.
"""

import json
import re
import tempfile
import uuid
from pathlib import Path
from threading import Thread

from fastapi import FastAPI, File, UploadFile, HTTPException, Request, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse

from mtg_packing_slip_organizer import (
    parse_packing_slip,
    fetch_colors_from_scryfall,
    extract_text_from_pdf,
    generate_html,
    fetch_scryfall_sets,
    get_set_sync_status,
)

app = FastAPI(title="MTG Packing Slip Organizer")

# Allow cross-origin requests from any WordPress site.
# You can restrict this to your specific domain once deployed, e.g.:
#   allow_origins=["https://yourdomain.com"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

MAX_FILE_SIZE = 10 * 1024 * 1024  # 10 MB

# In-memory store for processing jobs (job_id -> job state)
_jobs: dict[str, dict] = {}


@app.get("/", response_class=HTMLResponse)
def index():
    """Simple landing page with upload form (for standalone use)."""
    return """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>MTG Packing Slip Organizer</title>
        <style>
            * { box-sizing: border-box; }
            body {
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                margin: 0; padding: 40px;
                background: #1a1a2e; color: #eee;
                display: flex; flex-direction: column; align-items: center;
            }
            h1 { margin-bottom: 8px; }
            p.sub { color: #aaa; margin-top: 0; margin-bottom: 30px; }
            .file-slots {
                display: flex; gap: 14px; max-width: 620px; width: 100%;
            }
            .file-slot {
                flex: 1; border: 2px dashed #444; border-radius: 12px;
                padding: 24px 14px; text-align: center; cursor: pointer;
                transition: border-color 0.2s, background 0.2s;
                position: relative; min-height: 100px;
                display: flex; flex-direction: column;
                align-items: center; justify-content: center; gap: 6px;
            }
            .file-slot:hover { border-color: #6c63ff; background: rgba(108,99,255,0.05); }
            .file-slot.filled { border-style: solid; background: rgba(108,99,255,0.08); }
            .file-slot.filled .slot-label { display: none; }
            .file-slot:not(.filled) .slot-file,
            .file-slot:not(.filled) .slot-remove { display: none; }
            .file-slot input { display: none; }
            .slot-label { font-size: 0.9em; color: #888; }
            .slot-file { font-size: 0.85em; color: #6c63ff; word-break: break-all; font-weight: 600; }
            .slot-remove {
                position: absolute; top: 6px; right: 8px;
                color: #888; font-size: 1.2em; cursor: pointer;
                width: 22px; height: 22px; display: flex;
                align-items: center; justify-content: center;
                border-radius: 50%; transition: background 0.2s;
            }
            .slot-remove:hover { background: rgba(255,255,255,0.1); color: #ff6b6b; }
            .slot-group {
                position: absolute; bottom: 6px; right: 8px;
                font-size: 0.7em; font-weight: 700; color: #555;
                width: 20px; height: 20px; border-radius: 50%;
                display: flex; align-items: center; justify-content: center;
            }
            #slot0 .slot-group { background: rgba(74,158,255,0.3); color: #4a9eff; }
            #slot1 .slot-group { background: rgba(255,152,0,0.3); color: #ff9800; }
            #slot2 .slot-group { background: rgba(102,187,106,0.3); color: #66bb6a; }
            .file-slot:not(.filled) .slot-group { display: none; }
            button {
                margin-top: 20px; padding: 12px 40px;
                background: #6c63ff; color: #fff; border: none;
                border-radius: 8px; font-size: 1em; cursor: pointer;
                transition: background 0.2s;
            }
            button:hover { background: #5a52d5; }
            button:disabled { background: #444; cursor: not-allowed; }
            .status { margin-top: 20px; color: #aaa; min-height: 1.5em; }
            .error { color: #ff6b6b; }

            /* Processing progress */
            .processing { display: none; max-width: 500px; width: 100%; margin-top: 20px; }
            .processing.active { display: block; }
            .proc-bar-wrap {
                background: #16213e; border-radius: 6px; height: 24px;
                overflow: hidden; position: relative;
            }
            .proc-bar-fill {
                background: linear-gradient(90deg, #6c63ff, #48c6ef);
                height: 100%; width: 0%; transition: width 0.3s;
                border-radius: 6px;
            }
            .proc-bar-text {
                position: absolute; top: 0; left: 0; right: 0; bottom: 0;
                display: flex; align-items: center; justify-content: center;
                font-size: 0.8em; color: #fff; font-weight: 600;
            }
            .proc-detail {
                margin-top: 8px; font-size: 0.85em; color: #888;
                white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
            }

            /* Set status */
            .set-status {
                margin-top: 30px; padding: 16px 24px;
                background: #16213e; border-radius: 8px;
                max-width: 500px; width: 100%;
                font-size: 0.9em; color: #aaa;
            }
            .set-status-header {
                display: flex; justify-content: space-between; align-items: center;
            }
            .set-status .dot {
                display: inline-block; width: 8px; height: 8px;
                border-radius: 50%; margin-right: 8px;
            }
            .set-status .dot.ok { background: #2ecc71; }
            .set-status .dot.err { background: #e74c3c; }
            .set-status button {
                margin: 0; padding: 6px 14px;
                font-size: 0.85em; background: #2a2a4a;
            }
            .set-status button:hover { background: #3a3a5a; }
            .latest-sets {
                margin-top: 12px; padding-top: 12px;
                border-top: 1px solid #2a2a4a;
                font-size: 0.85em;
            }
            .latest-sets .label { color: #888; margin-bottom: 6px; }
            .latest-sets ul {
                list-style: none; margin: 0; padding: 0;
            }
            .latest-sets li {
                padding: 3px 0;
                display: flex; justify-content: space-between;
            }
            .latest-sets .set-name { color: #ccc; }
            .latest-sets .set-date { color: #666; font-size: 0.9em; }
        </style>
    </head>
    <body>
        <h1>MTG Packing Slip Organizer</h1>
        <p class="sub">Upload a TCGplayer packing slip PDF to generate an organized pull sheet</p>

        <div class="file-slots" id="fileSlots">
            <div class="file-slot" id="slot0" onclick="addFile(0)">
                <input type="file" accept=".pdf" onchange="handleSlotFile(0, this)">
                <div class="slot-label">PDF 1</div>
                <div class="slot-file"></div>
                <div class="slot-remove" onclick="removeFile(event, 0)">&times;</div>
                <div class="slot-group">A</div>
            </div>
            <div class="file-slot" id="slot1" onclick="addFile(1)">
                <input type="file" accept=".pdf" onchange="handleSlotFile(1, this)">
                <div class="slot-label">PDF 2 (optional)</div>
                <div class="slot-file"></div>
                <div class="slot-remove" onclick="removeFile(event, 1)">&times;</div>
                <div class="slot-group">B</div>
            </div>
            <div class="file-slot" id="slot2" onclick="addFile(2)">
                <input type="file" accept=".pdf" onchange="handleSlotFile(2, this)">
                <div class="slot-label">PDF 3 (optional)</div>
                <div class="slot-file"></div>
                <div class="slot-remove" onclick="removeFile(event, 2)">&times;</div>
                <div class="slot-group">C</div>
            </div>
        </div>
        <p class="sub" style="margin-top: 12px; font-size: 0.85em;">Upload 1 PDF for a single order, or up to 3 to merge into a multi-order pull sheet</p>

        <button id="submitBtn" disabled>Generate Pull Sheet</button>

        <div class="processing" id="processing">
            <div class="proc-bar-wrap">
                <div class="proc-bar-fill" id="procFill"></div>
                <div class="proc-bar-text" id="procText">Starting...</div>
            </div>
            <div class="proc-detail" id="procDetail"></div>
        </div>

        <div class="status" id="status"></div>

        <div class="set-status" id="setStatus">
            <div class="set-status-header">
                <span id="setInfo">Checking Scryfall set data...</span>
                <button id="refreshBtn" onclick="refreshSets()">Refresh Sets</button>
            </div>
            <div class="latest-sets" id="latestSets" style="display:none">
                <div class="label">Latest sets synced:</div>
                <ul id="latestSetsList"></ul>
            </div>
        </div>

        <script>
            const submitBtn = document.getElementById('submitBtn');
            const status = document.getElementById('status');
            const processing = document.getElementById('processing');
            const procFill = document.getElementById('procFill');
            const procText = document.getElementById('procText');
            const procDetail = document.getElementById('procDetail');

            // Multi-file slot management
            const slotFiles = [null, null, null];

            function addFile(idx) {
                const slot = document.getElementById('slot' + idx);
                slot.querySelector('input').click();
            }

            function handleSlotFile(idx, input) {
                const file = input.files[0];
                if (!file) return;
                if (!file.name.toLowerCase().endsWith('.pdf')) {
                    status.textContent = 'Please select a PDF file.';
                    status.className = 'status error';
                    return;
                }
                slotFiles[idx] = file;
                const slot = document.getElementById('slot' + idx);
                slot.classList.add('filled');
                slot.querySelector('.slot-file').textContent = file.name;
                status.textContent = '';
                status.className = 'status';
                updateSubmitState();
            }

            function removeFile(e, idx) {
                e.stopPropagation();
                slotFiles[idx] = null;
                const slot = document.getElementById('slot' + idx);
                slot.classList.remove('filled');
                slot.querySelector('.slot-file').textContent = '';
                slot.querySelector('input').value = '';
                updateSubmitState();
            }

            function updateSubmitState() {
                submitBtn.disabled = !slotFiles.some(f => f !== null);
            }

            // Drag and drop on individual slots
            document.querySelectorAll('.file-slot').forEach((slot, idx) => {
                slot.addEventListener('dragover', e => { e.preventDefault(); slot.style.borderColor = '#6c63ff'; });
                slot.addEventListener('dragleave', () => { slot.style.borderColor = ''; });
                slot.addEventListener('drop', e => {
                    e.preventDefault();
                    slot.style.borderColor = '';
                    if (e.dataTransfer.files.length) {
                        const input = slot.querySelector('input');
                        input.files = e.dataTransfer.files;
                        handleSlotFile(idx, input);
                    }
                });
            });

            // --- Set status ---
            async function loadSetStatus() {
                try {
                    const res = await fetch('/api/sets');
                    const data = await res.json();
                    const dot = data.cache_populated ? 'ok' : 'err';
                    const label = data.cache_populated
                        ? `<span class="dot ${dot}"></span>${data.sets_loaded} sets synced from Scryfall`
                        : `<span class="dot ${dot}"></span>Set sync failed — click Refresh`;
                    document.getElementById('setInfo').innerHTML = label;

                    // Render latest sets
                    const container = document.getElementById('latestSets');
                    const list = document.getElementById('latestSetsList');
                    if (data.latest_sets && data.latest_sets.length) {
                        list.innerHTML = data.latest_sets.map(s =>
                            `<li><span class="set-name">${s.name} <code style="color:#555">(${s.code})</code></span><span class="set-date">${s.released_at}</span></li>`
                        ).join('');
                        container.style.display = 'block';
                    } else {
                        container.style.display = 'none';
                    }
                } catch {
                    document.getElementById('setInfo').innerHTML =
                        '<span class="dot err"></span>Could not reach API';
                }
            }
            async function refreshSets() {
                const btn = document.getElementById('refreshBtn');
                btn.disabled = true;
                btn.textContent = 'Refreshing...';
                try {
                    await fetch('/api/sets/refresh', { method: 'POST' });
                    await loadSetStatus();
                } finally {
                    btn.disabled = false;
                    btn.textContent = 'Refresh Sets';
                }
            }
            loadSetStatus();

            // --- Upload + SSE progress ---
            submitBtn.addEventListener('click', async () => {
                const activeFiles = slotFiles.filter(f => f !== null);
                if (!activeFiles.length) return;
                submitBtn.disabled = true;
                status.textContent = '';
                status.className = 'status';

                // Show progress bar
                processing.classList.add('active');
                procFill.style.width = '0%';
                procText.textContent = activeFiles.length > 1 ? 'Uploading PDFs...' : 'Uploading PDF...';
                procDetail.textContent = '';

                // Step 1: Upload the files
                const form = new FormData();
                activeFiles.forEach(f => form.append('files', f));
                let jobId;

                try {
                    const uploadRes = await fetch('/api/parse', { method: 'POST', body: form });
                    if (!uploadRes.ok) {
                        const err = await uploadRes.json();
                        throw new Error(err.detail || 'Upload failed');
                    }
                    const uploadData = await uploadRes.json();
                    jobId = uploadData.job_id;
                    procText.textContent = `Parsing PDF (${uploadData.card_count} cards)...`;
                    procDetail.textContent = 'Looking up card data on Scryfall...';
                } catch (e) {
                    processing.classList.remove('active');
                    status.textContent = e.message;
                    status.className = 'status error';
                    submitBtn.disabled = false;
                    return;
                }

                // Step 2: Stream progress via SSE
                try {
                    const evtSource = new EventSource(`/api/parse/${jobId}/progress`);

                    evtSource.addEventListener('progress', (e) => {
                        const d = JSON.parse(e.data);
                        const pct = Math.round((d.current / d.total) * 100);
                        procFill.style.width = pct + '%';
                        procText.textContent = `Looking up cards: ${d.current} / ${d.total}`;
                        procDetail.textContent = `${d.card_name} — ${d.status}`;
                    });

                    evtSource.addEventListener('complete', (e) => {
                        evtSource.close();
                        const d = JSON.parse(e.data);

                        // Show done state briefly
                        procFill.style.width = '100%';
                        procText.textContent = 'Done!';
                        procDetail.textContent = 'Opening pull sheet...';

                        // Open result
                        const blob = new Blob([d.html], { type: 'text/html' });
                        window.open(URL.createObjectURL(blob), '_blank');

                        setTimeout(() => {
                            processing.classList.remove('active');
                            status.textContent = 'Pull sheet opened in a new tab.';
                            status.className = 'status';
                            submitBtn.disabled = false;
                        }, 800);
                    });

                    evtSource.addEventListener('error_event', (e) => {
                        evtSource.close();
                        const d = JSON.parse(e.data);
                        processing.classList.remove('active');
                        status.textContent = d.detail || 'Processing failed';
                        status.className = 'status error';
                        submitBtn.disabled = false;
                    });

                    evtSource.onerror = () => {
                        evtSource.close();
                        processing.classList.remove('active');
                        status.textContent = 'Connection lost during processing.';
                        status.className = 'status error';
                        submitBtn.disabled = false;
                    };
                } catch (e) {
                    processing.classList.remove('active');
                    status.textContent = e.message;
                    status.className = 'status error';
                    submitBtn.disabled = false;
                }
            });
        </script>
    </body>
    </html>
    """


@app.on_event("startup")
def startup_sync_sets():
    """Pre-fetch Scryfall set data on startup so the first request is fast."""
    fetch_scryfall_sets()


@app.get("/health")
def health():
    """Health check for deployment platforms."""
    return {"status": "ok"}


@app.get("/api/sets")
def sets_status():
    """Return set sync status and diagnostics."""
    return get_set_sync_status()


@app.post("/api/sets/refresh")
def refresh_sets():
    """Force a refresh of the Scryfall set data."""
    import mtg_packing_slip_organizer as mod
    mod._scryfall_set_cache = None
    mod._set_prefix_cache = None
    mod._scryfall_set_count = 0
    fetch_scryfall_sets()
    return get_set_sync_status()


@app.post("/api/parse")
async def parse_pdf(files: list[UploadFile] = File(...)):
    """Accept up to 3 packing slip PDFs, start processing, and return a job ID for progress tracking."""

    if len(files) > 3:
        raise HTTPException(status_code=400, detail="Maximum 3 PDFs supported.")

    group_labels = ['A', 'B', 'C']
    is_multi = len(files) > 1
    all_cards = []
    order_numbers = {}

    for idx, file in enumerate(files):
        if not file.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"File '{file.filename}' must be a PDF.")

        contents = await file.read()
        if len(contents) > MAX_FILE_SIZE:
            raise HTTPException(status_code=400, detail=f"File '{file.filename}' too large (max 10 MB).")

        # Write to a temp file so pdfplumber can read it
        tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        tmp.write(contents)
        tmp.close()
        tmp_path = tmp.name

        try:
            cards = parse_packing_slip(tmp_path)

            # Extract order number
            text = extract_text_from_pdf(tmp_path)
            order_match = re.search(r"Order\s*Number:\s*([A-Z0-9-]+)", text)
            order_num = order_match.group(1) if order_match else ""
        finally:
            Path(tmp_path).unlink(missing_ok=True)

        if not cards:
            raise HTTPException(
                status_code=422,
                detail=f"No cards found in '{file.filename}'. Please check the file format.",
            )

        # Assign order group for multi-order
        group = group_labels[idx] if is_multi else None
        if group:
            for card in cards:
                card.order_group = group
            order_numbers[group] = order_num
        else:
            order_numbers[''] = order_num

        all_cards.extend(cards)

    # Create a job for the slow Scryfall lookup phase
    job_id = uuid.uuid4().hex[:12]
    order_label = "Multi-Order Pull Sheet" if is_multi else order_numbers.get('', '')
    _jobs[job_id] = {
        "status": "processing",
        "cards": all_cards,
        "order_number": order_label,
        "order_numbers": order_numbers if is_multi else None,
        "progress": [],
        "result_html": None,
        "error": None,
    }

    # Run Scryfall lookups in a background thread
    def _process():
        job = _jobs[job_id]
        try:
            def on_progress(current, total, card_name, card_status):
                job["progress"].append({
                    "current": current,
                    "total": total,
                    "card_name": card_name,
                    "status": card_status,
                })

            fetch_colors_from_scryfall(job["cards"], on_progress=on_progress)
            html = generate_html(
                job["cards"],
                order_number=job["order_number"],
                order_numbers=job["order_numbers"],
            )
            job["result_html"] = html
            job["status"] = "complete"
        except Exception as e:
            job["error"] = str(e)
            job["status"] = "error"

    Thread(target=_process, daemon=True).start()

    return {"job_id": job_id, "card_count": len(all_cards), "order_count": len(files)}


@app.get("/api/parse/{job_id}/progress")
async def parse_progress(job_id: str, request: Request):
    """Stream processing progress via Server-Sent Events."""
    import asyncio

    if job_id not in _jobs:
        raise HTTPException(status_code=404, detail="Job not found.")

    async def event_stream():
        job = _jobs[job_id]
        sent = 0

        while True:
            # Check if client disconnected
            if await request.is_disconnected():
                break

            # Send any new progress events
            while sent < len(job["progress"]):
                evt = job["progress"][sent]
                yield f"event: progress\ndata: {json.dumps(evt)}\n\n"
                sent += 1

            # Check if done
            if job["status"] == "complete":
                yield f"event: complete\ndata: {json.dumps({'html': job['result_html']})}\n\n"
                # Clean up job
                del _jobs[job_id]
                break
            elif job["status"] == "error":
                yield f"event: error_event\ndata: {json.dumps({'detail': job['error']})}\n\n"
                del _jobs[job_id]
                break

            await asyncio.sleep(0.2)

    return StreamingResponse(
        event_stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
