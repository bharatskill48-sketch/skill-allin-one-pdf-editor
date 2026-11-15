
import * as pdfjsLib from 'pdfjs-dist';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import './style.css';

// PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  'pdfjs-dist/build/pdf.worker.min.mjs',
  import.meta.url
).toString();

// App state
const $ = selector => document.querySelector(selector);

let currentFile = null;
let pdfDoc = null;
let pdfBytes = null;
let pageNum = 1;
let scale = 1.15;
let imageFile = null;
let annotations = [];
let activeTool = 'select';
let drawing = null;

// Main application
const app = $('#app');

app.innerHTML = `
<aside class="sidebar">
  <div class="brand">
    <div class="brandmark">S</div>
    <div>
      <b>SKILL</b>
      <small>ALL IN ONE EDITOR</small>
    </div>
  </div>

  <button class="nav active" data-view="home">⌂ <span>Workspace</span></button>
  <button class="nav" data-view="pdf">▤ <span>PDF Editor</span></button>
  <button class="nav" data-view="image">▧ <span>Image Studio</span></button>
  <button class="nav" data-view="files">▦ <span>All Files</span></button>

  <div class="side-bottom">
    <div class="protag">PRO WORKSPACE</div>
    <small>Local file processing</small>
  </div>
</aside>

<main class="main">

<header>
  <div>
    <div class="eyebrow">DOCUMENT WORKSPACE</div>
    <h1 id="heading">Your files, one workspace.</h1>
  </div>

  <div class="head-actions">
    <button class="ghost" id="themeBtn">◐ Theme</button>
    <button class="primary" id="openBtn">＋ Open file</button>
    <input id="fileInput" type="file" hidden multiple>
  </div>
</header>

<section id="homeView" class="view">

  <div class="hero">
    <div>
      <span class="pill">ALL-IN-ONE DOCUMENT TOOLKIT</span>

      <h2>
        Work with documents<br>
        <em>without the busywork.</em>
      </h2>

      <p>Open, review, annotate and export your files.</p>

      <button class="whitebtn" id="heroOpen">
        Choose a file →
      </button>
    </div>

    <div class="hero-art">
      <div class="filetile pdf">PDF <strong>▤</strong></div>
      <div class="filetile cad">DWG <strong>⌗</strong></div>
      <div class="filetile img">PNG <strong>▧</strong></div>
    </div>
  </div>

  <div class="section-title">
    <div>
      <h3>Quick tools</h3>
      <p>Choose your workspace</p>
    </div>
  </div>

  <div class="tools-grid">

    <button class="tool-card" data-view="pdf">
      <span class="tool-icon purple">▤</span>
      <b>PDF Editor</b>
      <small>View, annotate and export PDF</small>
    </button>

    <button class="tool-card" data-view="image">
      <span class="tool-icon orange">▧</span>
      <b>Image Studio</b>
      <small>Preview, rotate and download images</small>
    </button>

    <button class="tool-card" data-view="files">
      <span class="tool-icon blue">▦</span>
      <b>Office & Files</b>
      <small>Preview documents and spreadsheets</small>
    </button>

    <button class="tool-card" id="convertCard">
      <span class="tool-icon green">⇄</span>
      <b>Convert to PDF</b>
      <small>Convert images and text into PDF</small>
    </button>

  </div>

  <div class="notice">
    <b>Private by design</b>
    <span>Files are processed locally in your browser.</span>
  </div>

</section>

<section id="pdfView" class="view hidden">

  <div class="editor-toolbar">

    <div class="filelabel" id="pdfFileLabel">
      No PDF selected
    </div>

    <div class="toolbar-group">
      <button data-tool="select" class="tool active">Select</button>
      <button data-tool="text" class="tool">＋ Text</button>
      <button data-tool="draw" class="tool">✎ Draw</button>
      <button id="clearMarks" class="tool">Clear marks</button>
    </div>

    <div class="toolbar-group">
      <button id="zoomOut" class="tool">−</button>
      <span id="zoomLabel">115%</span>
      <button id="zoomIn" class="tool">＋</button>
      <button id="downloadPdf" class="primary small" disabled>
        Export PDF
      </button>
    </div>

  </div>

  <div class="pdf-stage">

    <div id="pdfEmpty" class="empty">
      <div class="empty-icon">▤</div>
      <h3>Open a PDF to begin</h3>
      <p>View pages, add text notes and draw annotations.</p>
      <button class="primary" id="pdfChoose">Open PDF</button>
    </div>

    <div id="pdfCanvasWrap" class="canvas-wrap hidden">
      <canvas id="pdfCanvas"></canvas>
      <canvas id="markCanvas"></canvas>
    </div>

  </div>

  <div class="pagebar">
    <button id="prevPage" class="tool">← Previous</button>
    <span id="pageLabel">Page 0 / 0</span>
    <button id="nextPage" class="tool">Next →</button>
  </div>

</section>

<section id="imageView" class="view hidden">

  <div class="editor-toolbar">
    <div class="filelabel" id="imageLabel">No image selected</div>

    <div class="toolbar-group">
      <button id="rotateImage" class="tool">↻ Rotate</button>
      <button id="resetImage" class="tool">Reset</button>
      <button id="downloadImage" class="primary small" disabled>
        Download PNG
      </button>
    </div>
  </div>

  <div class="image-stage">

    <div id="imageEmpty" class="empty">
      <div class="empty-icon">▧</div>
      <h3>Open an image</h3>
      <p>PNG, JPG, WEBP and GIF preview.</p>
      <button class="primary" id="imageChoose">Open image</button>
    </div>

    <canvas id="imageCanvas" class="hidden"></canvas>

  </div>

</section>

<section id="filesView" class="view hidden">

  <div class="dropzone" id="dropzone">
    <div class="empty-icon">⇧</div>
    <h3>Drop a file here</h3>
    <p>Choose a file from your device.</p>
    <button class="primary" id="filesChoose">Browse files</button>
  </div>

  <div class="section-title">
    <div>
      <h3>Supported formats</h3>
      <p>Available file previews</p>
    </div>
  </div>

  <div class="format-grid">
    <div><b>PDF</b><small>View, annotate and export</small></div>
    <div><b>PNG / JPG</b><small>Image preview and rotate</small></div>
    <div><b>DOCX</b><small>Document preview</small></div>
    <div><b>XLSX / CSV</b><small>Spreadsheet preview</small></div>
    <div><b>TXT</b><small>Read text</small></div>
    <div><b>DWG / DXF</b><small>CAD engine required</small></div>
  </div>

  <div id="filePreview" class="file-preview hidden"></div>

</section>

<footer>
  SKILL ALL IN ONE PDF EDITOR
  <span>Web App · Version 1.0</span>
</footer>

</main>

<div id="toast" class="toast"></div>
`;

// Notification
function toast(message) {
  const element = $('#toast');

  element.textContent = message;
  element.classList.add('show');

  setTimeout(() => {
    element.classList.remove('show');
  }, 2800);
}

// Navigation
function showView(name) {

  document.querySelectorAll('.view').forEach(view => {
    view.classList.add('hidden');
  });

  $('#' + name + 'View').classList.remove('hidden');

  document.querySelectorAll('.nav').forEach(nav => {
    nav.classList.toggle('active', nav.dataset.view === name);
  });

  $('#heading').textContent = {
    home: 'Your files, one workspace.',
    pdf: 'PDF Editor',
    image: 'Image Studio',
    files: 'All Files & Formats'
  }[name];
}

document.querySelectorAll('[data-view]').forEach(button => {
  button.addEventListener('click', () => {
    showView(button.dataset.view);
  });
});

// Theme
$('#themeBtn').onclick = () => {
  document.body.classList.toggle('light');
};

// File selection
function chooseFile() {
  $('#fileInput').click();
}

$('#openBtn').onclick = chooseFile;
$('#heroOpen').onclick = chooseFile;
$('#pdfChoose').onclick = chooseFile;
$('#imageChoose').onclick = chooseFile;
$('#filesChoose').onclick = chooseFile;
$('#convertCard').onclick = chooseFile;

$('#fileInput').addEventListener('change', event => {

  [...event.target.files].forEach(handleFile);

  // Allow selecting the same file again
  event.target.value = '';
});

// File routing
async function handleFile(file) {

  currentFile = file;

  const ext = file.name.split('.').pop().toLowerCase();

  if (ext === 'pdf') {

    await openPdf(file);
    showView('pdf');

  } else if (
    ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp'].includes(ext)
  ) {

    openImage(file);
    showView('image');

  } else {

    showView('files');
    await previewOther(file, ext);

  }
}

// PDF OPEN
async function openPdf(file) {

  try {

    pdfBytes = await file.arrayBuffer();

    pdfDoc = await pdfjsLib.getDocument({
      data: pdfBytes.slice(0)
    }).promise;

    pageNum = 1;
    annotations = [];

    $('#pdfFileLabel').textContent = file.name;

    $('#pdfEmpty').classList.add('hidden');
    $('#pdfCanvasWrap').classList.remove('hidden');

    $('#downloadPdf').disabled = false;

    await renderPage();

    toast('PDF opened successfully');

  } catch (error) {

    console.error(error);
    toast('Could not open PDF. Check the file.');

  }
}

// PDF RENDER
async function renderPage() {

  if (!pdfDoc) return;

  const page = await pdfDoc.getPage(pageNum);

  const viewport = page.getViewport({
    scale
  });

  const canvas = $('#pdfCanvas');
  const marks = $('#markCanvas');

  canvas.width = viewport.width;
  canvas.height = viewport.height;

  marks.width = viewport.width;
  marks.height = viewport.height;

  canvas.style.width = viewport.width + 'px';
  canvas.style.height = viewport.height + 'px';

  marks.style.width = viewport.width + 'px';
  marks.style.height = viewport.height + 'px';

  await page.render({
    canvasContext: canvas.getContext('2d'),
    viewport
  }).promise;

  drawMarks();

  $('#pageLabel').textContent =
    `Page ${pageNum} / ${pdfDoc.numPages}`;

  $('#zoomLabel').textContent =
    Math.round(scale * 100) + '%';

  $('#prevPage').disabled = pageNum <= 1;
  $('#nextPage').disabled = pageNum >= pdfDoc.numPages;
}

// PDF PAGE NAVIGATION
$('#prevPage').onclick = () => {

  if (pdfDoc && pageNum > 1) {
    pageNum--;
    renderPage();
  }

};

$('#nextPage').onclick = () => {

  if (pdfDoc && pageNum < pdfDoc.numPages) {
    pageNum++;
    renderPage();
  }

};

// PDF ZOOM
$('#zoomIn').onclick = () => {

  scale = Math.min(2.5, scale + 0.15);

  if (pdfDoc) renderPage();

};

$('#zoomOut').onclick = () => {

  scale = Math.max(0.5, scale - 0.15);

  if (pdfDoc) renderPage();

};

// PDF TOOLS
document.querySelectorAll('[data-tool]').forEach(button => {

  button.onclick = () => {

    activeTool = button.dataset.tool;

    document.querySelectorAll('[data-tool]').forEach(item => {
      item.classList.toggle('active', item === button);
    });

    $('#markCanvas').style.cursor =
      activeTool === 'draw' ? 'crosshair' :
      activeTool === 'text' ? 'text' : 'default';

  };

});

// DRAW ANNOTATIONS
function drawMarks() {

  const canvas = $('#markCanvas');
  const ctx = canvas.getContext('2d');

  ctx.clearRect(0, 0, canvas.width, canvas.height);

  annotations
    .filter(item => item.page === pageNum)
    .forEach(item => {

      if (item.type === 'text') {

        ctx.fillStyle = '#d52d35';
        ctx.font = '18px Arial';
        ctx.fillText(item.text, item.x, item.y);

      } else {

        ctx.strokeStyle = '#e03c35';
        ctx.lineWidth = 3;

        ctx.beginPath();
        ctx.moveTo(item.x, item.y);
        ctx.lineTo(item.x2, item.y2);
        ctx.stroke();

      }

    });

}

// CLEAR CURRENT PAGE ANNOTATIONS
$('#clearMarks').onclick = () => {

  annotations = annotations.filter(item => {
    return item.page !== pageNum;
  });

  drawMarks();

};

// POINTER DOWN
$('#markCanvas').addEventListener('pointerdown', event => {

  if (activeTool === 'select') return;

  const rect = event.target.getBoundingClientRect();

  const x = (event.clientX - rect.left) *
    (event.target.width / rect.width);

  const y = (event.clientY - rect.top) *
    (event.target.height / rect.height);

  if (activeTool === 'text') {

    const text = prompt('Enter annotation text:');

    if (text) {

      annotations.push({
        type: 'text',
        page: pageNum,
        x,
        y,
        text
      });

      drawMarks();

    }

    return;
  }

  drawing = {
    type: 'line',
    page: pageNum,
    x,
    y,
    x2: x,
    y2: y
  };

  event.target.setPointerCapture(event.pointerId);

});

// POINTER MOVE
$('#markCanvas').addEventListener('pointermove', event => {

  if (!drawing) return;

  const rect = event.target.getBoundingClientRect();

  drawing.x2 = (event.clientX - rect.left) *
    (event.target.width / rect.width);

  drawing.y2 = (event.clientY - rect.top) *
    (event.target.height / rect.height);

  drawMarks();

  const ctx = event.target.getContext('2d');

  ctx.strokeStyle = '#e03c35';
  ctx.lineWidth = 3;

  ctx.beginPath();
  ctx.moveTo(drawing.x, drawing.y);
  ctx.lineTo(drawing.x2, drawing.y2);
  ctx.stroke();

});

// POINTER UP
$('#markCanvas').addEventListener('pointerup', () => {

  if (drawing) {

    annotations.push({
      ...drawing
    });

    drawing = null;

    drawMarks();

  }

});

$('#markCanvas').addEventListener('pointercancel', () => {
  drawing = null;
  drawMarks();
});

// PDF EXPORT
$('#downloadPdf').onclick = async () => {

  if (!pdfBytes) return;

  try {

    const doc = await PDFDocument.load(pdfBytes);

    const font = await doc.embedFont(
      StandardFonts.Helvetica
    );

    for (const item of annotations) {

      if (item.page > doc.getPageCount()) continue;

      const page = doc.getPage(item.page - 1);
      const size = page.getSize();

      const canvas = $('#markCanvas');

      const sx = size.width / canvas.width;
      const sy = size.height / canvas.height;

      if (item.type === 'text') {

        page.drawText(item.text, {

          x: item.x * sx,

          y: size.height - item.y * sy,

          size: 12,

          font,

          color: rgb(0.82, 0.12, 0.15)

        });

      } else {

        page.drawLine({

          start: {
            x: item.x * sx,
            y: size.height - item.y * sy
          },

          end: {
            x: item.x2 * sx,
            y: size.height - item.y2 * sy
          },

          thickness: 2,

          color: rgb(0.85, 0.15, 0.15)

        });

      }

    }

    const output = await doc.save();

    download(

      new Blob([output], {
        type: 'application/pdf'
      }),

      `edited-${currentFile.name}`

    );

    toast('Edited PDF exported');

  } catch (error) {

    console.error(error);
    toast('PDF export failed');

  }

};

// DOWNLOAD HELPER
function download(blob, name) {

  const url = URL.createObjectURL(blob);

  const link = document.createElement('a');

  link.href = url;
  link.download = name;

  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1500);

}

// IMAGE EDITOR
function openImage(file) {

  imageFile = file;

  const image = new Image();

  image.onload = () => {

    const canvas = $('#imageCanvas');
    const ctx = canvas.getContext('2d');

    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;

    ctx.drawImage(image, 0, 0);

    $('#imageEmpty').classList.add('hidden');
    canvas.classList.remove('hidden');

    $('#imageLabel').textContent = file.name;

    $('#downloadImage').disabled = false;

    URL.revokeObjectURL(image.src);

  };

  image.onerror = () => {
    toast('Image could not be opened');
  };

  image.src = URL.createObjectURL(file);

}

// IMAGE ROTATE
function rotateImage() {

  const canvas = $('#imageCanvas');

  const oldCanvas = document.createElement('canvas');

  oldCanvas.width = canvas.width;
  oldCanvas.height = canvas.height;

  oldCanvas.getContext('2d').drawImage(canvas, 0, 0);

  canvas.width = oldCanvas.height;
  canvas.height = oldCanvas.width;

  const ctx = canvas.getContext('2d');

  ctx.translate(
    canvas.width / 2,
    canvas.height / 2
  );

  ctx.rotate(Math.PI / 2);

  ctx.drawImage(
    oldCanvas,
    -oldCanvas.width / 2,
    -oldCanvas.height / 2
  );

}

$('#rotateImage').onclick = rotateImage;

// IMAGE RESET
$('#resetImage').onclick = () => {

  if (imageFile) {
    openImage(imageFile);
  }

};

// IMAGE DOWNLOAD
$('#downloadImage').onclick = () => {

  $('#imageCanvas').toBlob(blob => {

    if (blob) {
      download(blob, 'edited-image.png');
    }

  }, 'image/png');

};

// SAFE HTML
function escapeHtml(value) {

  return String(value).replace(/[&<>"']/g, char => ({

    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'

  })[char]);

}

// OFFICE FILE PREVIEW
async function previewOther(file, ext) {

  const box = $('#filePreview');

  box.classList.remove('hidden');

  box.innerHTML = `
    <h3>${escapeHtml(file.name)}</h3>
    <p>
      ${(file.size / 1024).toFixed(1)} KB ·
      ${ext.toUpperCase()}
    </p>
  `;

  try {

    if (ext === 'txt') {

      const text = await file.text();

      const pre = document.createElement('pre');

      pre.textContent = text.slice(0, 50000);

      box.appendChild(pre);

    }

    else if (ext === 'csv') {

      const workbook = XLSX.read(
        await file.text(),
        { type: 'string' }
      );

      renderSheet(workbook, box);

    }

    else if (ext === 'xlsx' || ext === 'xls') {

      const workbook = XLSX.read(
        await file.arrayBuffer()
      );

      renderSheet(workbook, box);

    }

    else if (ext === 'docx') {

      const result = await mammoth.convertToHtml({

        arrayBuffer: await file.arrayBuffer()

      });

      const article = document.createElement('article');

      article.className = 'docx';

      article.innerHTML = result.value;

      box.appendChild(article);

    }

    else if (ext === 'dwg' || ext === 'dxf') {

      const notice = document.createElement('div');

      notice.className = 'notice';

      notice.textContent =
        'CAD file detected. DWG/DXF preview and editing require a dedicated licensed CAD engine or secured backend integration.';

      box.appendChild(notice);

    }

    else {

      box.innerHTML += `
        <p>Preview is not available for this format yet.</p>
      `;

    }

  } catch (error) {

    console.error(error);

    box.innerHTML += `
      <p>Could not preview this file:
      ${escapeHtml(error.message)}</p>
    `;

  }

}

// SPREADSHEET PREVIEW
function renderSheet(workbook, box) {

  const sheet = workbook.Sheets[
    workbook.SheetNames[0]
  ];

  const table = document.createElement('div');

  table.innerHTML = XLSX.utils.sheet_to_html(sheet);

  box.appendChild(table);

}

// DRAG AND DROP
const dropzone = $('#dropzone');

dropzone.addEventListener('dragover', event => {

  event.preventDefault();

  dropzone.classList.add('drag');

});

dropzone.addEventListener('dragleave', () => {

  dropzone.classList.remove('drag');

});

dropzone.addEventListener('drop', event => {

  event.preventDefault();

  dropzone.classList.remove('drag');

  [...event.dataTransfer.files].forEach(handleFile);

});

// APPLICATION READY
toast('SKILL ALL IN ONE EDITOR ready');
  
