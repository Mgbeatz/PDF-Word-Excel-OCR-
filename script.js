/* script.js
 - Uses pdf.js (global: pdfjsLib),
 - Tesseract (global: Tesseract),
 - XLSX (global: XLSX),
 - htmlDocx (global: htmlDocx)
*/

pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.7.570/pdf.worker.min.js';

const dropArea = document.getElementById('drop-area');
const fileElem = document.getElementById('fileElem');
const jobsEl = document.getElementById('jobs');
const downloadAllBtn = document.getElementById('downloadAll');

let lastGenerated = { docxBlob: null, xlsxBlob: null };

['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt=>{
  dropArea.addEventListener(evt, preventDefaults, false);
});

function preventDefaults(e){
  e.preventDefault(); e.stopPropagation();
}
dropArea.addEventListener('dragover', ()=> dropArea.classList.add('dragover'));
dropArea.addEventListener('dragleave', ()=> dropArea.classList.remove('dragover'));
dropArea.addEventListener('drop', handleDrop);
fileElem.addEventListener('change', (e)=>handleFiles(e.target.files));

async function handleDrop(e){
  dropArea.classList.remove('dragover');
  const dt = e.dataTransfer;
  if(!dt) return;
  const files = Array.from(dt.files).filter(f=>f.type==='application/pdf');
  if(files.length===0) alert('Please drop PDF files only.');
  else handleFiles(files);
}

function handleFiles(files){
  Array.from(files).forEach(processPDFFile);
}

function createJobCard(filename){
  const card = document.createElement('div');
  card.className = 'job';
  card.innerHTML = `
    <strong>${escapeHtml(filename)}</strong>
    <div class="small status">Queued</div>
    <div class="progress-bar"><div class="progress-inner"></div></div>
    <div class="small result"></div>
  `;
  jobsEl.prepend(card);
  return card;
}

function updateCard(card, {status, progress, result}){
  if(status) card.querySelector('.status').textContent = status;
  if(typeof progress === 'number') card.querySelector('.progress-inner').style.width = `${progress}%`;
  if(result) card.querySelector('.result').innerHTML = result;
}

function escapeHtml(s){ return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }






async function processPDFFile(file) {
  const card = createJobCard(file.name);
  updateCard(card, {status: 'Reading PDF...', progress: 2});

  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({data: arrayBuffer});
  const pdf = await loadingTask.promise;
  const numPages = pdf.numPages;

  updateCard(card, {status:`Loaded — ${numPages} pages`, progress: 6});

  let fullText = '';
  const pagesText = [];

  for (let p = 1; p <= numPages; p++) {
    updateCard(card, {status: `Processing page ${p} / ${numPages}`, progress: 6 + Math.round((p - 1) / numPages * 20)});
    const page = await pdf.getPage(p);

    // Step 1 — Try extracting text normally
    const textContent = await page.getTextContent();
    const extracted = textContent.items.map(i => i.str).join(' ').trim();

    let pageText = '';
    if (extracted.length > 20) {
      // Found text — use it
      pageText = extracted;
      updateCard(card, {status: `Text extracted from page ${p}`, progress: 20 + Math.round(p / numPages * 40)});
    } else {
      // No text found — run OCR
      updateCard(card, {status: `Page ${p} seems scanned — running OCR...`, progress: 20 + Math.round(p / numPages * 40)});
      const viewport = page.getViewport({scale: 2.0});
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      const ctx = canvas.getContext('2d');
      await page.render({canvasContext: ctx, viewport}).promise;

      const dataUrl = canvas.toDataURL('image/png');
      try {
        const { data: { text } } = await Tesseract.recognize(dataUrl, 'eng', {
          logger: m => {
            if (m.status && m.progress) {
              updateCard(card, {status: `OCR: ${m.status} ${(m.progress*100).toFixed(0)}%`, progress: 40 + Math.round(p / numPages * 40 * m.progress)});
            }
          }
        });
        pageText = text.trim();
      } catch (err) {
        pageText = '[OCR failed]';
        console.error(err);
      }
    }

    pagesText.push(pageText);
    fullText += `<h3>Page ${p}</h3><pre>${escapeHtml(pageText)}</pre>`;
  }

  // Generate DOCX
  updateCard(card, {status: 'Generating DOCX...', progress: 95});
  const htmlForDoc = `<html><body><h1>${escapeHtml(file.name)}</h1>${fullText}</body></html>`;
  const docxBuf = htmlDocx.asBlob(htmlForDoc);
  const docxBlob = new Blob([docxBuf], {type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});

  // Generate XLSX
  const aoa = [['Page', 'Text']];
  pagesText.forEach((txt, idx) => aoa.push([idx + 1, txt]));
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Content');
  const xlsxArrayBuffer = XLSX.write(wb, {bookType:'xlsx', type:'array'});
  const xlsxBlob = new Blob([xlsxArrayBuffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

  // Download links
  const downloadLinks = `
    <div>
      <a download="${file.name.replace(/\.pdf$/i,'')}.docx" id="dl-docx">Download .docx</a> ·
      <a download="${file.name.replace(/\.pdf$/i,'')}.xlsx" id="dl-xlsx">Download .xlsx</a>
    </div>
  `;
  updateCard(card, {status: 'Done', progress: 100, result: downloadLinks});

  setTimeout(() => {
    card.querySelector('#dl-docx').href = URL.createObjectURL(docxBlob);
    card.querySelector('#dl-xlsx').href = URL.createObjectURL(xlsxBlob);

    lastGenerated = { docxBlob, xlsxBlob, name: file.name.replace(/\.pdf$/i,'') };
    downloadAllBtn.disabled = false;
  }, 50);
}


// Download last generated pair
downloadAllBtn.addEventListener('click', ()=>{
  if(!lastGenerated.docxBlob || !lastGenerated.xlsxBlob) return;
  const name = lastGenerated.name || 'output';
  triggerDownload(lastGenerated.docxBlob, `${name}.docx`);
  triggerDownload(lastGenerated.xlsxBlob, `${name}.xlsx`);
});

function triggerDownload(blob, filename){
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
}
