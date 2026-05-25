import { useState, useRef, useCallback } from 'react'

const initialForm = {
  ordemMissao: '',
  regFabViatura: '',
  motorista: '',
  odometroInicial: '',
  odometroFinal: '',
  destino: '',
  missao: '',
  peso: '',
  volume: '',
  localAbast: '',
  odometroAbast: '',
  qtde: '',
  data: '',
  tku: '',
}

export default function Home() {
  const [missions, setMissions] = useState([])
  const [form, setForm] = useState({ ...initialForm })
  const [period, setPeriod] = useState({ month: 'AGOSTO', start: '01/08/2025', end: '31/08/2025' })
  const tableRef = useRef(null)

  const handleChange = (field, value) => {
    setForm(prev => ({ ...prev, [field]: value }))
  }

  const calculateKM = () => {
    const ini = parseFloat(form.odometroInicial) || 0
    const fin = parseFloat(form.odometroFinal) || 0
    return Math.max(0, fin - ini)
  }

  const addMission = () => {
    const km = calculateKM()
    if (!form.ordemMissao || !form.regFabViatura || !form.motorista) return

    setMissions(prev => [...prev, {
      ...form,
      km,
      id: Date.now(),
    }])
    setForm({ ...initialForm })
  }

  const removeMission = (id) => {
    setMissions(prev => prev.filter(m => m.id !== id))
  }

  const clearAll = () => {
    setMissions([])
  }

  const totals = missions.reduce((acc, m) => ({
    km: acc.km + (m.km || 0),
    peso: acc.peso + (parseFloat(m.peso) || 0),
    volume: acc.volume + (parseFloat(m.volume) || 0),
    qtde: acc.qtde + (parseFloat(m.qtde) || 0),
  }), { km: 0, peso: 0, volume: 0, qtde: 0 })

  const exportToExcel = useCallback(async () => {
    const ExcelJS = (await import('exceljs')).default

    const wb = new ExcelJS.Workbook()
    wb.creator = 'CELOG'
    wb.created = new Date()

    const ws = wb.addWorksheet('Misses')

    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    }

    const yellowFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC000' } }
    const headerFont = { bold: true, size: 10, name: 'Calibri' }
    const dataFont = { size: 10, name: 'Calibri' }
    const boldFont = { bold: true, size: 10, name: 'Calibri' }

    const titleRow = ws.addRow([`VIAGENS MISSOES CELOG (${period.month} / ${period.start} a ${period.end})`])
    ws.mergeCells(`A${titleRow.number}:O${titleRow.number}`)
    const titleCell = titleRow.getCell(1)
    titleCell.font = { bold: true, size: 14, name: 'Calibri' }
    titleCell.alignment = { horizontal: 'center' }
    titleRow.height = 30

    ws.addRow([])

    const headers = [
      'Ordem de misso', 'REG FAB VIATURA', 'MOTORISTA',
      'ODOMETRO INICIAL', 'ODOMETRO FINAL', 'KM',
      'DESTINO', 'MISSO', 'PESO (KG)', 'VOLUME (m)',
      'LOCAL DE ABAST.', 'ODMETRO', 'QTDE (LTS)', 'DATA', 'TKU',
    ]

    ws.columns = headers.map((h, i) => {
      const widths = [20, 18, 18, 16, 16, 12, 30, 35, 12, 12, 20, 14, 12, 14, 10]
      return { width: widths[i] || 15 }
    })

    const headerRow = ws.addRow(headers)
    headerRow.eachCell((cell) => {
      cell.fill = yellowFill
      cell.font = headerFont
      cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true }
      cell.border = borderStyle
    })
    headerRow.height = 30

    missions.forEach(m => {
      const row = ws.addRow([
        m.ordemMissao, m.regFabViatura, m.motorista,
        m.odometroInicial, m.odometroFinal, m.km,
        m.destino, m.missao, m.peso, m.volume,
        m.localAbast, m.odometroAbast, m.qtde, m.data, m.tku,
      ])
      row.eachCell((cell) => {
        cell.font = dataFont
        cell.alignment = { vertical: 'center', wrapText: true }
        cell.border = borderStyle
      })
    })

    if (missions.length > 0) {
      ws.addRow([])
      const totalRow = ws.addRow([
        'TOTAL', '', '', '', '', totals.km, '', '', totals.peso, totals.volume,
        '', '', totals.qtde, '', '',
      ])
      totalRow.eachCell((cell) => {
        cell.font = boldFont
        cell.border = borderStyle
        if (cell.col === 1) cell.alignment = { horizontal: 'right' }
      })
      ws.mergeCells(`A${totalRow.number}:B${totalRow.number}`)
    }

    const buf = await wb.xlsx.writeBuffer()
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `VIAGENS_MISSOES_CELOG_${period.month}_${period.start.replace(/\//g, '')}_a_${period.end.replace(/\//g, '')}.xlsx`
    a.click()
    URL.revokeObjectURL(url)
  }, [missions, period, totals])

  const exportToPDF = useCallback(async () => {
    const html2canvas = (await import('html2canvas')).default
    const { jsPDF } = await import('jspdf')

    const element = tableRef.current
    if (!element) return

    const wrapper = element.querySelector('.table-wrapper')
    if (wrapper) wrapper.style.overflow = 'visible'

    const canvas = await html2canvas(element, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      width: element.scrollWidth,
      height: element.scrollHeight,
      windowWidth: element.scrollWidth,
    })

    if (wrapper) wrapper.style.overflow = ''

    const imgData = canvas.toDataURL('image/png')
    const pdf = new jsPDF('l', 'mm', 'a4')
    const landscapeWidth = 297
    const landscapeHeight = 210
    const margin = 10
    const pdfImgWidth = landscapeWidth - 2 * margin
    const pdfImgHeight = (canvas.height * pdfImgWidth) / canvas.width

    let heightLeft = pdfImgHeight
    let position = margin

    pdf.addImage(imgData, 'PNG', margin, position, pdfImgWidth, pdfImgHeight)
    heightLeft -= landscapeHeight

    while (heightLeft > 0) {
      position = margin - (pdfImgHeight - heightLeft)
      pdf.addPage()
      pdf.addImage(imgData, 'PNG', margin, position, pdfImgWidth, pdfImgHeight)
      heightLeft -= landscapeHeight
    }

    pdf.save(`VIAGENS_MISSOES_CELOG_${period.month}.pdf`)
  }, [period])

  return (
    <div className="container">
      <header className="header">
        <h1>VIAGENS MISSÕES CELOG</h1>
        <div className="period-row">
          <label>MÊS:</label>
          <input
            type="text"
            value={period.month}
            onChange={e => setPeriod(p => ({ ...p, month: e.target.value }))}
            className="period-input"
          />
          <label>PERÍODO:</label>
          <input
            type="text"
            value={period.start}
            onChange={e => setPeriod(p => ({ ...p, start: e.target.value }))}
            className="period-input date-input"
          />
          <span>a</span>
          <input
            type="text"
            value={period.end}
            onChange={e => setPeriod(p => ({ ...p, end: e.target.value }))}
            className="period-input date-input"
          />
        </div>
      </header>

      <section className="form-section">
        <h2>Cadastrar Missão</h2>
        <div className="form-grid">
          <div className="form-group">
            <label>Ordem de Missão</label>
            <input value={form.ordemMissao} onChange={e => handleChange('ordemMissao', e.target.value)} />
          </div>
          <div className="form-group">
            <label>REG FAB VIATURA</label>
            <input value={form.regFabViatura} onChange={e => handleChange('regFabViatura', e.target.value)} />
          </div>
          <div className="form-group">
            <label>MOTORISTA</label>
            <input value={form.motorista} onChange={e => handleChange('motorista', e.target.value)} />
          </div>
          <div className="form-group">
            <label>ODÔMETRO INICIAL</label>
            <input type="number" value={form.odometroInicial} onChange={e => handleChange('odometroInicial', e.target.value)} />
          </div>
          <div className="form-group">
            <label>ODÔMETRO FINAL</label>
            <input type="number" value={form.odometroFinal} onChange={e => handleChange('odometroFinal', e.target.value)} />
          </div>
          <div className="form-group">
            <label>KM</label>
            <input value={calculateKM()} readOnly className="calculated-field" />
          </div>
          <div className="form-group form-group-wide">
            <label>DESTINO</label>
            <input value={form.destino} onChange={e => handleChange('destino', e.target.value)} />
          </div>
          <div className="form-group form-group-wide">
            <label>MISSÃO</label>
            <input value={form.missao} onChange={e => handleChange('missao', e.target.value)} />
          </div>
          <div className="form-group">
            <label>PESO (KG)</label>
            <input type="number" step="0.01" value={form.peso} onChange={e => handleChange('peso', e.target.value)} />
          </div>
          <div className="form-group">
            <label>VOLUME (m³)</label>
            <input type="number" step="0.001" value={form.volume} onChange={e => handleChange('volume', e.target.value)} />
          </div>
          <div className="form-group">
            <label>LOCAL DE ABAST.</label>
            <input value={form.localAbast} onChange={e => handleChange('localAbast', e.target.value)} />
          </div>
          <div className="form-group">
            <label>ODÔMETRO</label>
            <input type="number" value={form.odometroAbast} onChange={e => handleChange('odometroAbast', e.target.value)} />
          </div>
          <div className="form-group">
            <label>QTDE (LTS)</label>
            <input type="number" step="0.1" value={form.qtde} onChange={e => handleChange('qtde', e.target.value)} />
          </div>
          <div className="form-group">
            <label>DATA</label>
            <input type="text" placeholder="dd/mm/aaaa" value={form.data} onChange={e => handleChange('data', e.target.value)} />
          </div>
          <div className="form-group">
            <label>TKU</label>
            <input value={form.tku} onChange={e => handleChange('tku', e.target.value)} />
          </div>
        </div>
        <div className="form-buttons">
          <button onClick={addMission} className="btn btn-primary">Adicionar Missão</button>
          <button onClick={() => setForm({ ...initialForm })} className="btn btn-secondary">Limpar Formulário</button>
        </div>
      </section>

      <section className="table-section" ref={tableRef}>
        <h2>{`Missões Cadastradas (${missions.length})`}</h2>
        {missions.length === 0 ? (
          <p className="empty-msg">Nenhuma missão cadastrada. Adicione missões pelo formulário acima.</p>
        ) : (
          <div className="table-wrapper">
            <table className="missions-table">
              <thead>
                <tr>
                  <th>Ordem de Missão</th>
                  <th>REG FAB VIATURA</th>
                  <th>MOTORISTA</th>
                  <th>ODÔMETRO INICIAL</th>
                  <th>ODÔMETRO FINAL</th>
                  <th>KM</th>
                  <th>DESTINO</th>
                  <th>MISSÃO</th>
                  <th>PESO (KG)</th>
                  <th>VOLUME (m³)</th>
                  <th>LOCAL DE ABAST.</th>
                  <th>ODÔMETRO</th>
                  <th>QTDE (LTS)</th>
                  <th>DATA</th>
                  <th>TKU</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {missions.map(m => (
                  <tr key={m.id}>
                    <td data-label="Ordem de Missão">{m.ordemMissao}</td>
                    <td data-label="REG FAB VIATURA">{m.regFabViatura}</td>
                    <td data-label="MOTORISTA">{m.motorista}</td>
                    <td data-label="OD. INICIAL" className="num">{m.odometroInicial}</td>
                    <td data-label="OD. FINAL" className="num">{m.odometroFinal}</td>
                    <td data-label="KM" className="num">{m.km}</td>
                    <td data-label="DESTINO">{m.destino}</td>
                    <td data-label="MISSÃO">{m.missao}</td>
                    <td data-label="PESO (KG)" className="num">{m.peso}</td>
                    <td data-label="VOLUME (m³)" className="num">{m.volume}</td>
                    <td data-label="LOCAL ABAST.">{m.localAbast}</td>
                    <td data-label="ODÔMETRO" className="num">{m.odometroAbast}</td>
                    <td data-label="QTDE (LTS)" className="num">{m.qtde}</td>
                    <td data-label="DATA">{m.data}</td>
                    <td data-label="TKU">{m.tku}</td>
                    <td data-label="">
                      <button onClick={() => removeMission(m.id)} className="btn-remove">✕</button>
                    </td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan="2" data-label="TOTAL"><strong>TOTAL ODÔMETRO</strong></td>
                  <td data-label=""></td>
                  <td data-label=""></td>
                  <td data-label=""></td>
                  <td data-label="KM" className="num"><strong>{totals.km}</strong></td>
                  <td data-label=""></td>
                  <td data-label="TOTAL"><strong>TOTAL PESO x CUBAGEM</strong></td>
                  <td data-label="PESO" className="num"><strong>{totals.peso.toFixed(2)}</strong></td>
                  <td data-label="VOLUME" className="num"><strong>{totals.volume.toFixed(3)}</strong></td>
                  <td data-label=""></td>
                  <td data-label=""></td>
                  <td data-label="QTDE" className="num"><strong>{totals.qtde.toFixed(1)}</strong></td>
                  <td data-label=""></td>
                  <td data-label=""></td>
                  <td data-label=""></td>
                </tr>
              </tfoot>
            </table>
          </div>
        )}
      </section>

      <section className="export-section">
        <h2>Exportar</h2>
        <div className="export-buttons">
          <button onClick={exportToExcel} className="btn btn-excel">
            Baixar Excel
          </button>
          <button onClick={exportToPDF} className="btn btn-pdf">
            Baixar PDF
          </button>
          <button onClick={clearAll} className="btn btn-danger">
            Limpar Tudo
          </button>
        </div>
      </section>
    </div>
  )
}
