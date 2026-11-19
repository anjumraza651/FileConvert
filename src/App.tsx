import { useRef, useState } from 'react';
import { PDFDocument } from 'pdf-lib';
import * as pdfjsLib from 'pdfjs-dist';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import mammoth from 'mammoth';
import PptxGenJS from 'pptxgenjs';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { FileText, FileImage, Files, FileSpreadsheet, Presentation, Upload, Download, X, ShieldCheck, Sparkles, LoaderCircle, CheckCircle2, ArrowRight } from 'lucide-react';

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).toString();

type ToolId = 'word-pdf' | 'pdf-word' | 'ppt-pdf' | 'pdf-ppt' | 'excel-pdf' | 'pdf-excel' | 'word-ppt' | 'images-pdf' | 'merge' | 'split' | 'pdf-images';
type Tool = { id: ToolId; title: string; description: string; accept: string; multiple?: boolean; group: string; output: string };
type Result = { name: string; url: string };

const tools: Tool[] = [
  { id: 'word-pdf', title: 'Word to PDF', description: 'Convert DOCX documents into PDFs.', accept: '.docx', group: 'Office to PDF', output: 'PDF' },
  { id: 'ppt-pdf', title: 'PowerPoint to PDF', description: 'Export PPTX slide text as a PDF.', accept: '.pptx', group: 'Office to PDF', output: 'PDF' },
  { id: 'excel-pdf', title: 'Excel to PDF', description: 'Export workbook sheets as a PDF.', accept: '.xlsx,.xls,.csv', group: 'Office to PDF', output: 'PDF' },
  { id: 'pdf-word', title: 'PDF to Word', description: 'Extract PDF text into an editable DOCX.', accept: '.pdf', group: 'PDF to Office', output: 'DOCX' },
  { id: 'pdf-ppt', title: 'PDF to PowerPoint', description: 'Create one text slide per PDF page.', accept: '.pdf', group: 'PDF to Office', output: 'PPTX' },
  { id: 'pdf-excel', title: 'PDF to Excel', description: 'Extract PDF text into a worksheet.', accept: '.pdf', group: 'PDF to Office', output: 'XLSX' },
  { id: 'word-ppt', title: 'Word to PowerPoint', description: 'Turn document paragraphs into slides.', accept: '.docx', group: 'Office conversions', output: 'PPTX' },
  { id: 'images-pdf', title: 'Images to PDF', description: 'Combine JPG and PNG images into a PDF.', accept: 'image/jpeg,image/png', multiple: true, group: 'PDF tools', output: 'PDF' },
  { id: 'merge', title: 'Merge PDFs', description: 'Combine multiple PDF documents.', accept: '.pdf', multiple: true, group: 'PDF tools', output: 'PDF' },
  { id: 'split', title: 'Split PDF', description: 'Export each PDF page separately.', accept: '.pdf', group: 'PDF tools', output: 'PDF pages' },
  { id: 'pdf-images', title: 'PDF to images', description: 'Export pages as PNG images.', accept: '.pdf', group: 'PDF tools', output: 'PNG' }
];

const isPdf = (file: File) => file.name.toLowerCase().endsWith('.pdf') || file.type === 'application/pdf';
const isDocx = (file: File) => file.name.toLowerCase().endsWith('.docx');
const isPptx = (file: File) => file.name.toLowerCase().endsWith('.pptx');
const isExcel = (file: File) => /\.(xlsx|xls|csv)$/i.test(file.name);
const lines = (text: string) => text.split(/\r?\n/).map(value => value.trim()).filter(Boolean);

async function extractPdfText(file: File): Promise<string[]> {
  const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  const pages: string[] = [];
  for (let index = 1; index <= pdf.numPages; index++) {
    const page = await pdf.getPage(index);
    const content = await page.getTextContent();
    pages.push(content.items.map(item => 'str' in item ? item.str : '').join(' ').trim());
  }
  return pages;
}

async function textPdf(pages: string[]): Promise<Blob> {
  const { jsPDF } = await import('jspdf');
  const pdf = new jsPDF();
  pages.forEach((text, pageIndex) => {
    if (pageIndex) pdf.addPage();
    let y = 18;
    const wrapped = pdf.splitTextToSize(text || ' ', 175);
    wrapped.forEach((line: string) => {
      if (y > 280) { pdf.addPage(); y = 18; }
      pdf.text(line, 17, y);
      y += 7;
    });
  });
  return pdf.output('blob');
}

async function makeDocx(paragraphs: string[]): Promise<Blob> {
  const document = new Document({ sections: [{ children: paragraphs.map(text => new Paragraph({ children: [new TextRun(text)] })) }] });
  return Packer.toBlob(document);
}

async function makePptx(pages: string[]): Promise<Blob> {
  const presentation = new PptxGenJS();
  presentation.layout = 'LAYOUT_WIDE';
  pages.forEach((text, index) => {
    const slide = presentation.addSlide();
    slide.addText(text || ' ', { x: 0.6, y: 0.5, w: 12, h: 6.3, fontFace: 'Aptos', fontSize: 20, valign: 'top', margin: 0.15 });
    slide.addText(String(index + 1), { x: 12.1, y: 7.05, w: 0.5, h: 0.2, fontSize: 9, color: '888888' });
  });
  return presentation.write({ outputType: 'blob' }) as Promise<Blob>;
}

async function extractSlides(file: File): Promise<string[]> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const names = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name)).sort((a, b) => Number(a.match(/slide(\d+)/)?.[1]) - Number(b.match(/slide(\d+)/)?.[1]));
  const pages: string[] = [];
  for (const name of names) {
    const xml = await zip.file(name)!.async('text');
    const parsed = new DOMParser().parseFromString(xml, 'application/xml');
    pages.push(Array.from(parsed.getElementsByTagName('a:t')).map(node => node.textContent || '').join(' '));
  }
  return pages;
}

function App() {
  const [selected, setSelected] = useState<ToolId>('word-pdf');
  const [files, setFiles] = useState<File[]>([]);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [results, setResults] = useState<Result[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);
  const current = tools.find(tool => tool.id === selected)!;

  function clearResults() {
    results.forEach(item => URL.revokeObjectURL(item.url));
    setResults([]);
    setMessage('');
  }

  function chooseTool(id: ToolId) {
    clearResults();
    setSelected(id);
    setFiles([]);
    if (inputRef.current) inputRef.current.value = '';
  }

  function addFiles(list: FileList | null) {
    if (!list) return;
    const incoming = Array.from(list);
    const valid = incoming.filter(file => {
      if (selected === 'word-pdf' || selected === 'word-ppt') return isDocx(file);
      if (selected === 'ppt-pdf') return isPptx(file);
      if (selected === 'excel-pdf') return isExcel(file);
      if (selected === 'images-pdf') return ['image/jpeg', 'image/png'].includes(file.type);
      return isPdf(file);
    });
    if (valid.length !== incoming.length) setMessage('One or more files are not supported for this conversion.');
    else setMessage('');
    setFiles(previous => selected === 'merge' || selected === 'images-pdf' ? [...previous, ...valid] : valid.slice(0, 1));
    clearResults();
  }

  function result(blob: Blob, name: string): Result {
    return { name, url: URL.createObjectURL(blob) };
  }

  async function processFiles() {
    if (!files.length || busy) return;
    setBusy(true);
    clearResults();
    try {
      const file = files[0];
      let output: Result[] = [];
      if (selected === 'word-pdf' || selected === 'word-ppt') {
        if (!isDocx(file)) throw new Error('Please upload a DOCX Word document.');
        const extracted = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
        const paragraphs = lines(extracted.value);
        if (!paragraphs.length) throw new Error('No readable text was found in this Word document.');
        const blob = selected === 'word-pdf' ? await textPdf(paragraphs) : await makePptx(paragraphs);
        output = [result(blob, file.name.replace(/\.docx$/i, selected === 'word-pdf' ? '.pdf' : '.pptx'))];
      } else if (selected === 'ppt-pdf') {
        if (!isPptx(file)) throw new Error('Please upload a PPTX presentation.');
        const pages = await extractSlides(file);
        if (!pages.length) throw new Error('No readable slide text was found.');
        output = [result(await textPdf(pages), file.name.replace(/\.pptx$/i, '.pdf'))];
      } else if (selected === 'excel-pdf') {
        if (!isExcel(file)) throw new Error('Please upload an Excel or CSV file.');
        const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array' });
        const pages = workbook.SheetNames.map(name => name + '\n' + XLSX.utils.sheet_to_csv(workbook.Sheets[name]));
        output = [result(await textPdf(pages), file.name.replace(/\.(xlsx|xls|csv)$/i, '.pdf'))];
      } else if (selected === 'pdf-word') {
        const pages = await extractPdfText(file);
        const paragraphs = pages.flatMap((page, index) => [index ? '' : '', ...lines(page)]);
        output = [result(await makeDocx(paragraphs), file.name.replace(/\.pdf$/i, '.docx'))];
      } else if (selected === 'pdf-ppt') {
        const pages = await extractPdfText(file);
        output = [result(await makePptx(pages.map((page, index) => 'Page ' + (index + 1) + '\n' + page)), file.name.replace(/\.pdf$/i, '.pptx'))];
      } else if (selected === 'pdf-excel') {
        const pages = await extractPdfText(file);
        const rows = pages.flatMap((page, index) => [['Page ' + (index + 1)], ...lines(page).map(line => [line])]);
        const sheet = XLSX.utils.aoa_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, sheet, 'Extracted text');
        const bytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        output = [result(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), file.name.replace(/\.pdf$/i, '.xlsx'))];
      } else if (selected === 'images-pdf') {
        const pdf = await PDFDocument.create();
        for (const imageFile of files) {
          const bytes = await imageFile.arrayBuffer();
          const image = imageFile.type === 'image/png' ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes);
          const scale = Math.min(595.28 / image.width, 841.89 / image.height, 1);
          const page = pdf.addPage([image.width * scale, image.height * scale]);
          page.drawImage(image, { x: 0, y: 0, width: image.width * scale, height: image.height * scale });
        }
        output = [result(new Blob([await pdf.save()], { type: 'application/pdf' }), 'images-converted.pdf')];
      } else if (selected === 'merge') {
        if (files.length < 2) throw new Error('Add at least two PDF files to merge.');
        const merged = await PDFDocument.create();
        for (const item of files) {
          const source = await PDFDocument.load(await item.arrayBuffer());
          const copied = await merged.copyPages(source, source.getPageIndices());
          copied.forEach(page => merged.addPage(page));
        }
        output = [result(new Blob([await merged.save()], { type: 'application/pdf' }), 'merged-document.pdf')];
      } else if (selected === 'split') {
        const source = await PDFDocument.load(await file.arrayBuffer());
        for (let index = 0; index < source.getPageCount(); index++) {
          const part = await PDFDocument.create();
          const [page] = await part.copyPages(source, [index]);
          part.addPage(page);
          output.push(result(new Blob([await part.save()], { type: 'application/pdf' }), 'page-' + (index + 1) + '.pdf'));
        }
      } else {
        const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
          const page = await pdf.getPage(pageNumber);
          const viewport = page.getViewport({ scale: 1.5 });
          const canvas = document.createElement('canvas');
          const context = canvas.getContext('2d');
          if (!context) throw new Error('Could not create an image canvas.');
          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);
          await page.render({ canvasContext: context, viewport }).promise;
          const blob = await new Promise<Blob>((resolve, reject) => canvas.toBlob(value => value ? resolve(value) : reject(new Error('Could not export page image.')), 'image/png'));
          output.push(result(blob, 'page-' + pageNumber + '.png'));
          canvas.width = 0;
          canvas.height = 0;
        }
      }
      setResults(output);
      setMessage('Done — ' + output.length + ' file' + (output.length === 1 ? '' : 's') + ' ready to download.');
    } catch (error) {
      setMessage(error instanceof Error ? error.message : 'Conversion failed. Please try another file.');
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className='app-shell'>
      <header className='topbar'>
        <a className='brand' href='#' onClick={event => event.preventDefault()}>
          <span className='brand-mark'><FileText size={21} /></span>
          <span>File<span className='brand-light'>Convert</span></span>
        </a>
        <div className='privacy'><ShieldCheck size={15} /> Files stay on your device</div>
      </header>
      <main>
        <section className='hero'>
          <div className='eyebrow'><Sparkles size={14} /> ALL-IN-ONE FILE CONVERTER</div>
          <h1>Free Online File Converter -<br /><span>PDF, Word & PowerPoint</span></h1>
          <p>
  Convert PDF to Word, PDF to PowerPoint, Word to PDF, Excel to PDF,
  PowerPoint to PDF and images to PDF online for free.
  Merge PDFs, split PDF pages and convert your documents easily.
  No account required. Your files stay private on your device.
</p> </section>
        <section className='workspace'>
          <div className='tool-directory'>
            {['Office to PDF', 'PDF to Office', 'Office conversions', 'PDF tools'].map(group => (
              <div className='tool-group' key={group}>
                <h3>{group}</h3>
                <div className='tool-grid'>
                  {tools.filter(item => item.group === group).map(item => (
                    <button key={item.id} className={selected === item.id ? 'tool-card chosen' : 'tool-card'} onClick={() => chooseTool(item.id)}>
                      <span className='tool-icon'>{item.id.includes('excel') ? <FileSpreadsheet size={19} /> : item.id.includes('ppt') ? <Presentation size={19} /> : item.id.includes('images') ? <FileImage size={19} /> : item.id === 'merge' || item.id === 'split' ? <Files size={19} /> : <FileText size={19} />}</span>
                      <strong>{item.title}</strong><small>{item.description}</small>
                    </button>
                  ))}
                </div>
              </div>
            ))}
          </div>
          <div className='panel'>
            <div className='panel-heading'><div><h2>{current.title}</h2><p>{current.description}</p></div><span className='step-pill'>01 <span>·</span> UPLOAD</span></div>
            <button className='dropzone' onClick={() => inputRef.current?.click()} onDragOver={event => event.preventDefault()} onDrop={event => { event.preventDefault(); addFiles(event.dataTransfer.files); }}>
              <span className='upload-icon'><Upload size={23} /></span>
              <strong>Choose {current.multiple ? 'files' : 'a file'} or drag and drop</strong>
              <span>Accepted: {current.accept}</span><small>{current.multiple ? 'You can select multiple files' : 'One file per conversion'} · Your files stay private</small>
            </button>
            <input ref={inputRef} className='visually-hidden' type='file' accept={current.accept} multiple={!!current.multiple} onChange={event => addFiles(event.target.files)} />
            {files.length > 0 && <div className='file-list'>
              <div className='file-list-head'><strong>{files.length} {files.length === 1 ? 'file' : 'files'} selected</strong><button className='text-button' onClick={() => { setFiles([]); clearResults(); }}>Clear all</button></div>
              {files.map((file, index) => <div className='file-row' key={file.name + file.size + index}>
                <span className='file-icon'><FileText size={17} /></span><span className='file-name'><strong>{file.name}</strong><small>{(file.size / 1024 / 1024).toFixed(2)} MB</small></span>
                <button className='remove-button' aria-label={'Remove ' + file.name} onClick={() => { setFiles(previous => previous.filter((_, i) => i !== index)); clearResults(); }}><X size={16} /></button>
              </div>)}
            </div>}
            <div className='action-row'><span className='file-hint'>{current.output} output · Browser-based conversion</span>
              <button className='primary-button' disabled={!files.length || busy || (selected === 'merge' && files.length < 2)} onClick={processFiles}>
                {busy ? <><LoaderCircle className='spin' size={17} /> Working...</> : <>Convert now <ArrowRight size={17} /></>}
              </button>
            </div>
            {message && <div className={results.length ? 'notice success' : 'notice'}>{results.length > 0 && <CheckCircle2 size={17} />}{message}</div>}
            {results.length > 0 && <div className='result-list'>{results.map(item => <a className='download-button' href={item.url} download={item.name} key={item.url}><Download size={17} /> Download {item.name}</a>)}</div>}
          </div>
        </section>
        <section className='seo-content' aria-labelledby='seo-content-title'>
          <h2 id='seo-content-title'>Free Online File Conversion Tools</h2>
          <p>
            FileConvert helps you convert documents online in your browser.
            Convert Word documents to PDF, PDF files to Word, PowerPoint
            presentations to PDF, PDF files to PowerPoint, and Excel
            spreadsheets to PDF. You can also convert Word documents to
            PowerPoint, turn JPG and PNG images into PDFs, merge PDF files,
            split PDF pages, and export PDF pages as PNG images.
          </p>

          <h2>Simple and Private File Conversion</h2>
          <p>
            Choose a conversion tool, select a supported file, and download
            the converted result. FileConvert processes files in your browser,
            so your selected documents stay on your device. No account is
            required.
          </p>

          <h2>Supported Conversion Tools</h2>
          <ul>
            <li>Word to PDF and Word to PowerPoint</li>
            <li>PDF to Word, PowerPoint, and Excel</li>
            <li>PowerPoint to PDF and Excel to PDF</li>
            <li>JPG and PNG images to PDF</li>
            <li>Merge PDFs, split PDFs, and PDF to PNG images</li>
          </ul>
        </section>
        <section className='trust-row'>
          
          <div><ShieldCheck size={19} /><span><strong>Private by design</strong><small>Files are processed in this browser.</small></span></div>
          <div><Sparkles size={19} /><span><strong>Many formats</strong><small>Office documents, PDFs and images.</small></span></div>
          <div><CheckCircle2 size={19} /><span><strong>No account needed</strong><small>Convert and download instantly.</small></span></div>
        </section>
        <p className='conversion-note'>Text-based conversions preserve extracted text, not the original layout, fonts, images or complex formatting. Scanned PDFs need OCR and are not supported by text extraction tools.</p>
      </main>
      <footer>FileConvert <span>·</span> Your documents stay yours.</footer>
    </div>
  );
}

export default App;
