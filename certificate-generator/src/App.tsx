import { useEffect, useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { Upload, FileSpreadsheet, Image as ImageIcon, Download, Users, Award, Settings2, CheckCircle2, MousePointer2, RefreshCw } from 'lucide-react'

type Person = { name: string; [key: string]: string }
const fontCatalog: Record<string, { family: string; fallback: string }> = {
  Georgia: { family: 'Georgia', fallback: 'serif' },
  Arial: { family: 'Arial', fallback: 'sans-serif' },
  'Times New Roman': { family: 'Times New Roman', fallback: 'serif' },
  Verdana: { family: 'Verdana', fallback: 'sans-serif' },
  'Trebuchet MS': { family: 'Trebuchet MS', fallback: 'sans-serif' },
  'Courier New': { family: 'Courier New', fallback: 'monospace' },
  'Segoe UI': { family: 'Segoe UI', fallback: 'sans-serif' },
  Tahoma: { family: 'Tahoma', fallback: 'sans-serif' },
  Impact: { family: 'Impact', fallback: 'sans-serif' }
}
const fonts = Object.keys(fontCatalog)

async function ensureFontReady(fontName: string) {
  const target = fontCatalog[fontName]?.family ?? fontName
  if (typeof document !== 'undefined' && 'fonts' in document) {
    const fontCheck = `16px "${target}"`
    if (!document.fonts.check(fontCheck)) {
      await document.fonts.ready
    }
  }
}

function toCanvasFont(fontName: string, fontSize: number, bold: boolean) {
  const config = fontCatalog[fontName] ?? { family: fontName, fallback: 'sans-serif' }
  return `${bold ? 'bold ' : ''}${fontSize}px "${config.family}", ${config.fallback}`
}

export default function App() {
  const [template, setTemplate] = useState<string>('')
  const [templateName, setTemplateName] = useState('')
  const [people, setPeople] = useState<Person[]>([])
  const [columns, setColumns] = useState<string[]>([])
  const [selectedColumn, setSelectedColumn] = useState('name')
  const [fontSize, setFontSize] = useState(44)
  const [fontFamily, setFontFamily] = useState('Georgia')
  const [fontColor, setFontColor] = useState('#172554')
  const [bold, setBold] = useState(true)
  const [x, setX] = useState(50)
  const [y, setY] = useState(55)
  const [previewIndex, setPreviewIndex] = useState(0)
  const [generating, setGenerating] = useState(false)
  const [generated, setGenerated] = useState(0)
  const [notice, setNotice] = useState('')
  const [imageDimensions, setImageDimensions] = useState({ width: 1200, height: 850 })
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const imageRef = useRef<HTMLImageElement | null>(null)
  const templateInputRef = useRef<HTMLInputElement>(null)
  const excelInputRef = useRef<HTMLInputElement>(null)

  const previewName = people[previewIndex]?.[selectedColumn] || people[previewIndex]?.name || 'Participant Name'
  const ready = !!template && people.length > 0

  useEffect(() => {
    if (!template) return
    const img = new Image()
    img.onload = () => {
      imageRef.current = img
      setImageDimensions({ width: img.naturalWidth, height: img.naturalHeight })
      void draw(img, previewName)
    }
    img.src = template
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [template])

  useEffect(() => {
    if (imageRef.current) {
      void draw(imageRef.current, previewName)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [previewName, x, y, fontSize, fontFamily, fontColor, bold, imageDimensions])

  async function draw(img: HTMLImageElement, name: string) {
    await ensureFontReady(fontFamily)
    const canvas = canvasRef.current
    if (!canvas) return
    const ctx = canvas.getContext('2d')
    if (!ctx) return
    canvas.width = img.naturalWidth || 1200
    canvas.height = img.naturalHeight || 850
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
    ctx.save()
    ctx.fillStyle = fontColor
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.font = toCanvasFont(fontFamily, fontSize, bold)
    const maxWidth = canvas.width * 0.84
    let text = name || 'Participant Name'
    while (ctx.measureText(text).width > maxWidth && ctx.font.includes('px')) {
      const sizeMatch = ctx.font.match(/(\d+)px/)
      const size = sizeMatch ? Number(sizeMatch[1]) : fontSize
      if (size <= 12) break
      ctx.font = toCanvasFont(fontFamily, size - 1, bold)
    }
    ctx.fillText(text, canvas.width * x / 100, canvas.height * y / 100, maxWidth)
    ctx.restore()
  }

  async function handleTemplate(file?: File) {
    if (!file) return
    if (!file.type.startsWith('image/')) {
      setNotice('Please upload a PNG or JPG image certificate template.')
      return
    }
    setTemplateName(file.name)
    setTemplate(await fileToDataUrl(file))
    setNotice('Certificate template loaded.')
    setGenerated(0)
  }

  async function handleExcel(file?: File) {
    if (!file) return
    try {
      const workbook = new ExcelJS.Workbook()
      await workbook.xlsx.load(await file.arrayBuffer())
      const sheet = workbook.worksheets[0]
      if (!sheet) throw new Error('No worksheet found.')
      const headerRow = sheet.getRow(1)
      const headers: string[] = []
      headerRow.eachCell((cell, col) => { headers[col - 1] = String(cell.value ?? `Column ${col}`).trim() })
      const cleanHeaders = headers.map((h, i) => h || `Column ${i + 1}`)
      const rows: Person[] = []
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return
        const item: Person = { name: '' }
        cleanHeaders.forEach((header, index) => {
          const value = row.getCell(index + 1).text.trim()
          item[header] = value
        })
        const first = Object.values(item).find(v => v.trim())
        if (first) item.name = first
        rows.push(item)
      })
      if (!rows.length) throw new Error('No participant rows found. Use the first row for column headings.')
      setColumns(cleanHeaders)
      const likelyName = cleanHeaders.find(h => /name|participant|student/i.test(h)) || cleanHeaders[0]
      setSelectedColumn(likelyName)
      setPeople(rows)
      setPreviewIndex(0)
      setGenerated(0)
      setNotice(`Imported ${rows.length} participants from ${file.name}.`)
    } catch (err) {
      setNotice(err instanceof Error ? err.message : 'Could not read this Excel file.')
    }
  }

  function onCanvasClick(event: React.MouseEvent<HTMLCanvasElement>) {
    const rect = event.currentTarget.getBoundingClientRect()
    setX(Math.round(((event.clientX - rect.left) / rect.width) * 100))
    setY(Math.round(((event.clientY - rect.top) / rect.height) * 100))
  }

  async function renderCertificate(name: string): Promise<Blob> {
    const img = imageRef.current
    if (!img) throw new Error('Upload a certificate image first.')
    const canvas = document.createElement('canvas')
    canvas.width = img.naturalWidth
    canvas.height = img.naturalHeight
    const ctx = canvas.getContext('2d')!
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
    await ensureFontReady(fontFamily)
    ctx.fillStyle = fontColor
    ctx.textAlign = 'center'
    ctx.textBaseline = 'middle'
    ctx.font = toCanvasFont(fontFamily, fontSize, bold)
    ctx.fillText(name, canvas.width * x / 100, canvas.height * y / 100, canvas.width * .84)
    return new Promise((resolve, reject) => canvas.toBlob(blob => blob ? resolve(blob) : reject(new Error('Image export failed.')), 'image/png'))
  }

  async function generateAll() {
    if (!ready) return
    setGenerating(true)
    setNotice('')
    try {
      const zip = new JSZip()
      for (let i = 0; i < people.length; i++) {
        const name = people[i][selectedColumn] || people[i].name || `Participant ${i + 1}`
        const blob = await renderCertificate(name)
        zip.file(`${safeFileName(name)}.png`, blob)
      }
      const archive = await zip.generateAsync({ type: 'blob' })
      downloadBlob(archive, 'certificates.zip')
      setGenerated(people.length)
      setNotice(`Done! ${people.length} certificates were generated and downloaded as a ZIP.`)
    } catch (err) {
      setNotice(err instanceof Error ? err.message : 'Certificate generation failed.')
    } finally {
      setGenerating(false)
    }
  }

  async function downloadOne() {
    if (!previewName || !imageRef.current) return
    const blob = await renderCertificate(previewName)
    downloadBlob(blob, `${safeFileName(previewName)}.png`)
  }

  const stats = useMemo(() => ({ participants: people.length, generated }), [people.length, generated])

  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><div className="brand-icon"><Award size={22}/></div><div><strong>Certify</strong><span>Certificate Studio</span></div></div>
      <div className="side-label">WORKSPACE</div>
      <div className="nav-item active"><Settings2 size={18}/> Certificate Builder</div>
      <div className="side-note"><div className="note-icon"><CheckCircle2 size={17}/></div><div><b>Private by design</b><p>Your files stay in this browser. Nothing is uploaded to a server.</p></div></div>
      <div className="sidebar-bottom">CERTIFY STUDIO <span>v1.0</span></div>
    </aside>
    <main className="main">
      <header className="topbar"><div><div className="eyebrow">WORKSPACE / BUILDER</div><h1>Certificate Generator</h1><p>Create personalized certificates in just a few clicks.</p></div><div className="top-pill"><span className="live-dot"/> Browser-based</div></header>
      <section className="stats-grid">
        <Stat icon={<Users size={19}/>} label="Participants" value={stats.participants.toString()} />
        <Stat icon={<Award size={19}/>} label="Certificates generated" value={stats.generated.toString()} />
        <Stat icon={<FileSpreadsheet size={19}/>} label="Excel status" value={people.length ? 'Imported' : 'Not uploaded'} small />
      </section>
      <section className="workspace-grid">
        <div className="left-column">
          <div className="panel">
            <div className="panel-heading"><div className="step">01</div><div><h2>Upload your files</h2><p>Start with a certificate design and a participant list.</p></div></div>
            <div className="upload-grid">
              <div className={`upload-card ${template ? 'uploaded' : ''}`} onClick={() => templateInputRef.current?.click()}>
                <input ref={templateInputRef} type="file" accept="image/png,image/jpeg,image/webp" hidden onChange={e => handleTemplate(e.target.files?.[0])}/>
                <div className="upload-icon purple"><ImageIcon size={21}/></div><b>Certificate template</b><span>{templateName || 'PNG or JPG image'}</span><button className="outline-btn"><Upload size={15}/>{template ? 'Replace template' : 'Choose image'}</button>
              </div>
              <div className={`upload-card ${people.length ? 'uploaded' : ''}`} onClick={() => excelInputRef.current?.click()}>
                <input ref={excelInputRef} type="file" accept=".xlsx,.xls" hidden onChange={e => handleExcel(e.target.files?.[0])}/>
                <div className="upload-icon green"><FileSpreadsheet size={21}/></div><b>Participant Excel</b><span>{people.length ? `${people.length} participants loaded` : 'Excel workbook (.xlsx)'}</span><button className="outline-btn"><Upload size={15}/>{people.length ? 'Replace Excel' : 'Choose Excel'}</button>
              </div>
            </div>
            {people.length > 0 && <div className="column-select"><label>Name column</label><select value={selectedColumn} onChange={e => {setSelectedColumn(e.target.value);setGenerated(0)}}>{columns.map(c => <option key={c} value={c}>{c}</option>)}</select><span>{people.length} rows imported</span></div>}
          </div>
          <div className="panel preview-panel">
            <div className="panel-heading"><div className="step">02</div><div><h2>Preview & position</h2><p>Click anywhere on the certificate to place the name.</p></div></div>
            <div className="preview-toolbar"><span><MousePointer2 size={15}/> Click on the preview to set name position</span>{people.length > 0 && <div className="person-switch"><button disabled={previewIndex===0} onClick={()=>setPreviewIndex(i=>Math.max(0,i-1))}>‹</button><span>{previewIndex+1} / {people.length}</span><button disabled={previewIndex>=people.length-1} onClick={()=>setPreviewIndex(i=>Math.min(people.length-1,i+1))}>›</button></div>}</div>
            <div className="canvas-wrap">
              {template ? <canvas ref={canvasRef} onClick={onCanvasClick} aria-label="Certificate preview; click to position name"/> : <div className="empty-preview"><div className="empty-art"><Award size={42}/></div><b>Your certificate preview appears here</b><span>Upload a PNG or JPG template to get started.</span></div>}
            </div>
            {template && <div className="preview-footer"><span><span className="tiny-dot"/> Live preview</span><span>{imageDimensions.width} × {imageDimensions.height} px</span></div>}
          </div>
        </div>
        <div className="right-column">
          <div className="panel settings-panel">
            <div className="panel-heading"><div className="step">03</div><div><h2>Name styling</h2><p>Make the name match your design.</p></div></div>
            <div className="field"><label>Font family</label><select value={fontFamily} onChange={e=>setFontFamily(e.target.value)}>{fonts.map(f=><option key={f}>{f}</option>)}</select></div>
            <div className="field"><div className="label-row"><label>Font size</label><strong>{fontSize}px</strong></div><input type="range" min="12" max="100" value={fontSize} onChange={e=>setFontSize(Number(e.target.value))}/></div>
            <div className="field"><label>Text color</label><div className="color-row"><input type="color" value={fontColor} onChange={e=>setFontColor(e.target.value)}/><span>{fontColor.toUpperCase()}</span><div className="swatches">{['#172554','#111827','#7C2D12','#92400E','#FFFFFF'].map(c=><button key={c} style={{background:c}} aria-label={`Set color ${c}`} onClick={()=>setFontColor(c)}/>)}</div></div></div>
            <label className="toggle-row"><span><b>Bold text</b><small>Use a heavier font weight</small></span><input type="checkbox" checked={bold} onChange={e=>setBold(e.target.checked)}/></label>
            <div className="position-box"><div className="position-title"><b>Name position</b><span>Adjust with sliders or click preview</span></div><div className="field"><div className="label-row"><label>Horizontal</label><strong>{x}%</strong></div><input type="range" min="0" max="100" value={x} onChange={e=>setX(Number(e.target.value))}/></div><div className="field"><div className="label-row"><label>Vertical</label><strong>{y}%</strong></div><input type="range" min="0" max="100" value={y} onChange={e=>setY(Number(e.target.value))}/></div><button className="reset-btn" onClick={()=>{setX(50);setY(55)}}><RefreshCw size={14}/> Reset position</button></div>
          </div>
          <div className="panel generate-panel"><div className="generate-top"><div className="generate-icon"><Award size={22}/></div><div><h2>Ready to generate?</h2><p>{ready ? `${people.length} personalized certificates are ready.` : 'Upload both files to enable generation.'}</p></div></div><button className="generate-btn" disabled={!ready || generating} onClick={generateAll}>{generating ? <><RefreshCw className="spin" size={17}/> Generating…</> : <><Download size={17}/> Generate & download ZIP</>}</button><button className="download-one" disabled={!template || !people.length} onClick={downloadOne}>Download preview certificate</button>{notice && <div className={`notice ${notice.startsWith('Done') ? 'success' : ''}`}>{notice}</div>}<div className="format-note"><CheckCircle2 size={15}/> Exports high-resolution PNG files</div></div>
          {people.length > 0 && <div className="panel people-panel"><div className="people-heading"><h3>Participants</h3><span>{people.length}</span></div><div className="people-list">{people.slice(0,5).map((p,i)=><button key={i} className={i===previewIndex?'person-row selected':'person-row'} onClick={()=>setPreviewIndex(i)}><span className="avatar">{(p[selectedColumn]||p.name||'?').charAt(0).toUpperCase()}</span><span>{p[selectedColumn]||p.name||`Participant ${i+1}`}</span><span className="row-num">{String(i+1).padStart(2,'0')}</span></button>)}{people.length>5&&<div className="more-people">+ {people.length-5} more participants</div>}</div></div>}
        </div>
      </section>
      <footer>Made for event organizers <span>•</span> Files are processed locally in your browser</footer>
    </main>
  </div>
}
function Stat({icon,label,value,small=false}:{icon:React.ReactNode,label:string,value:string,small?:boolean}){return <div className="stat-card"><div className="stat-icon">{icon}</div><div><span>{label}</span><strong className={small?'small-value':''}>{value}</strong></div></div>}
function fileToDataUrl(file:File){return new Promise<string>((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(String(reader.result));reader.onerror=()=>reject(new Error('Could not read image.'));reader.readAsDataURL(file)})}
function safeFileName(name:string){return name.trim().replace(/[<>:"/\\|?*\x00-\x1F]/g,'_').replace(/\s+/g,'_').slice(0,100)||'certificate'}
function downloadBlob(blob:Blob,name:string){const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url)}
