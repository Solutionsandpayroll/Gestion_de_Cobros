import { useState, useEffect } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

function App() {
  const [isHelpExpanded, setIsHelpExpanded] = useState(false)
  const [excelFile, setExcelFile] = useState(null)
  const [dragActive, setDragActive] = useState(false)
  const [clientes, setClientes] = useState([])
  const [clientesSeleccionados, setClientesSeleccionados] = useState(new Set())
  const [clientesError, setClientesError] = useState('')
  const [busquedaCliente, setBusquedaCliente] = useState('')
  const [excelBuffer, setExcelBuffer] = useState(null)
  const [generando, setGenerando] = useState(false)

  useEffect(() => {
    if (!excelFile) {
      setClientes([])
      setClientesSeleccionados(new Set())
      setClientesError('')
      setExcelBuffer(null)
      return
    }
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array', cellStyles: true, sheetStubs: true })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]

        console.log('=== DEBUG EXCEL ===')
        console.log('Hoja activa:', workbook.SheetNames[0])
        console.log('Rango reportado:', sheet['!ref'])

        // Escanear TODAS las claves del sheet para encontrar celdas en fila 7
        // (independiente del rango reportado, que puede ser incorrecto por celdas combinadas)
        const row7Entries = Object.keys(sheet)
          .filter(key => !key.startsWith('!'))
          .map(key => {
            try { return { key, decoded: XLSX.utils.decode_cell(key), cell: sheet[key] } }
            catch { return null }
          })
          .filter(entry => entry && entry.decoded.r === 6)  // fila 7 = índice 6

        console.log('--- Celdas encontradas en fila 7 ---')
        row7Entries.forEach(({ key, cell }) => console.log(`  ${key}:`, JSON.stringify(cell.v), '| tipo:', cell.t))

        // Buscar la celda que diga "Cliente"
        const clienteEntry = row7Entries.find(({ cell }) => String(cell.v ?? '').trim() === 'Cliente')

        if (!clienteEntry) {
          setClientesError('No se encontró "Cliente" en la fila 7 del archivo.')
          setClientes([])
          console.log('Valores encontrados en fila 7:', row7Entries.map(e => e.cell.v))
          return
        }

        const clienteCol = clienteEntry.decoded.c
        console.log('Columna "Cliente" encontrada:', clienteEntry.key, '→ índice col:', clienteCol)

        // Leer todos los valores en esa columna desde fila 8 en adelante
        const clientesSet = new Set()
        const allRowsInCol = Object.keys(sheet)
          .filter(key => !key.startsWith('!'))
          .map(key => {
            try { return { decoded: XLSX.utils.decode_cell(key), cell: sheet[key] } }
            catch { return null }
          })
          .filter(entry => entry && entry.decoded.c === clienteCol && entry.decoded.r >= 7)
          .sort((a, b) => a.decoded.r - b.decoded.r)

        for (const { cell } of allRowsInCol) {
          const val = String(cell.v ?? '').trim()
          if (!val) break
          clientesSet.add(val)
        }

        console.log('Clientes únicos encontrados:', clientesSet.size, [...clientesSet].slice(0, 5))

        if (clientesSet.size === 0) {
          setClientesError('La columna "Cliente" no tiene datos a partir de la fila 8.')
          setClientes([])
          return
        }

        const lista = [...clientesSet].sort()
        setClientes(lista)
        setClientesSeleccionados(new Set(lista))
        setExcelBuffer(e.target.result)
        setClientesError('')
      } catch (err) {
        console.error('Error procesando Excel:', err)
        setClientesError('Error al leer el archivo. Verifica que sea un Excel válido.')
        setClientes([])
      }
    }
    reader.readAsArrayBuffer(excelFile)
  }, [excelFile])

  const CAMPOS = ['Cliente', 'Documento', 'Fecha vencimiento', 'Vencido 1 a 30', 'Vencido 31 a 60', 'Vencido 61 a 90', 'Vencido más de 91', 'Saldo por vencer']

  const generarFormato = async () => {
    if (!excelBuffer || clientesSeleccionados.size === 0) return
    setGenerando(true)
    try {
      // --- Leer SIGO con xlsx ---
      const sigoWb = XLSX.read(new Uint8Array(excelBuffer), { type: 'array' })
      const sigoSheet = sigoWb.Sheets[sigoWb.SheetNames[0]]

      const sigoEntries = Object.keys(sigoSheet)
        .filter(k => !k.startsWith('!'))
        .map(k => { try { return { decoded: XLSX.utils.decode_cell(k), cell: sigoSheet[k] } } catch { return null } })
        .filter(Boolean)

      const sigoColMap = {}
      for (const entry of sigoEntries.filter(e => e.decoded.r === 6)) {
        const val = String(entry.cell.v ?? '').trim()
        if (CAMPOS.includes(val)) sigoColMap[val] = entry.decoded.c
      }

      const sigoClienteCol = sigoColMap['Cliente']
      const maxRow = Math.max(...sigoEntries.map(e => e.decoded.r))
      const dataRows = []
      for (let r = 7; r <= maxRow; r++) {
        const clienteCell = sigoSheet[XLSX.utils.encode_cell({ r, c: sigoClienteCol })]
        if (!clienteCell || !clienteCell.v) break
        if (!clientesSeleccionados.has(String(clienteCell.v).trim())) continue
        const row = {}
        for (const campo of CAMPOS) {
          if (sigoColMap[campo] === undefined) continue
          const cell = sigoSheet[XLSX.utils.encode_cell({ r, c: sigoColMap[campo] })]
          row[campo] = cell ? cell.v : null
        }
        dataRows.push(row)
      }

      // --- Cargar plantilla con ExcelJS (preserva estilos, anchos, fuentes) ---
      const templateRes = await fetch('/Plantilla - Cuentas por cobrar detallada.xlsx')
      if (!templateRes.ok) throw new Error('No se pudo cargar la plantilla')

      // Corregir shared formulas con DOMParser (más robusto que regex para manejar
      // cualquier orden de atributos, citas simples/dobles y namespaces)
      const JSZip = (await import('jszip')).default
      const rawBuffer = await templateRes.arrayBuffer()
      const zip = await JSZip.loadAsync(rawBuffer)

      const xmlParser = new DOMParser()
      const xmlSerializer = new XMLSerializer()

      for (const [path, file] of Object.entries(zip.files)) {
        if (/^xl\/worksheets\/sheet\d+\.xml$/.test(path)) {
          const xml = await file.async('string')
          const doc = xmlParser.parseFromString(xml, 'application/xml')
          const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
          const fElements = Array.from(doc.getElementsByTagNameNS(ns, 'f'))
          for (const f of fElements) {
            if (f.getAttribute('t') !== 'shared') continue
            if (f.textContent.trim()) {
              // Master → fórmula normal
              f.removeAttribute('t')
              f.removeAttribute('ref')
              f.removeAttribute('si')
            } else {
              // Clone → eliminar (la celda mantiene su valor cacheado en <v>)
              f.parentNode.removeChild(f)
            }
          }
          zip.file(path, xmlSerializer.serializeToString(doc))
        }
      }

      const fixedBuffer = await zip.generateAsync({ type: 'arraybuffer' })
      const ExcelJS = (await import('exceljs')).default
      const templateWb = new ExcelJS.Workbook()
      await templateWb.xlsx.load(fixedBuffer)
      const ws = templateWb.worksheets[0]

      // Mapear encabezados de fila 1 en la plantilla
      const colMap = {}
      ws.getRow(1).eachCell({ includeEmpty: false }, (cell, col) => {
        const val = String(cell.value ?? '').trim()
        if (CAMPOS.includes(val)) colMap[val] = col
      })

      // Capturar estilos de referencia desde fila 2 (si existe) para aplicar a nuevas filas
      const refStyles = {}
      for (const campo of CAMPOS) {
        if (!colMap[campo]) continue
        const refCell = ws.getRow(2).getCell(colMap[campo])
        if (refCell && refCell.style) {
          refStyles[colMap[campo]] = JSON.parse(JSON.stringify(refCell.style))
        }
      }

      // Capturar filas fijas 7-23 ANTES de limpiar (con todos sus valores y estilos)
      const fijasCaptured = []
      for (let r = 7; r <= 23; r++) {
        const wsRow = ws.getRow(r)
        const cells = []
        wsRow.eachCell({ includeEmpty: true }, (cell, col) => {
          cells.push({
            col,
            value: cell.value,
            style: cell.style ? JSON.parse(JSON.stringify(cell.style)) : {},
            numFmt: cell.numFmt,
          })
        })
        fijasCaptured.push({ height: wsRow.height, cells })
      }

      // Limpiar filas de datos existentes en la plantilla (desde fila 2)
      const lastRow = ws.rowCount
      for (let r = lastRow; r >= 2; r--) ws.spliceRows(r, 1)

      // Escribir datos manteniendo estilos de referencia
      dataRows.forEach((rowData, i) => {
        const excelRow = ws.getRow(i + 2)
        for (const campo of CAMPOS) {
          if (!colMap[campo]) continue
          const cell = excelRow.getCell(colMap[campo])
          if (refStyles[colMap[campo]]) cell.style = JSON.parse(JSON.stringify(refStyles[colMap[campo]]))
          cell.value = rowData[campo] ?? null
        }
        excelRow.commit()
      })

      // Agregar fila vacía de separación y luego las filas fijas al final
      const firstFijaRow = 2 + dataRows.length + 1
      const separador = ws.getRow(firstFijaRow - 1)
      separador.commit()

      fijasCaptured.forEach((fija, i) => {
        const excelRow = ws.getRow(firstFijaRow + i)
        if (fija.height) excelRow.height = fija.height
        fija.cells.forEach(({ col, value, style, numFmt }) => {
          const cell = excelRow.getCell(col)
          cell.value = value
          if (style) cell.style = JSON.parse(JSON.stringify(style))
          if (numFmt) cell.numFmt = numFmt
        })
        excelRow.commit()
      })

      // Descargar
      const buffer = await templateWb.xlsx.writeBuffer()
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'Cuentas por cobrar.xlsx'
      a.click()
      URL.revokeObjectURL(url)

    } catch (err) {
      console.error('Error generando formato:', err)
      alert('Ocurrió un error al generar el archivo: ' + err.message)
    } finally {
      setGenerando(false)
    }
  }

  return (
    <div className="app">
      {/* Header Corporativo Solutions & Payroll */}
      <header className="header">
        <div className="container">
          <div className="header-content">
            <div className="logo-container">
              <div className="logo">
                <img 
                  src="/Logo syp.png" 
                  alt="Solutions & Payroll Logo" 
                  width="60" 
                  height="60"
                />
              </div>
              <div className="header-text">
                <h1>Solutions & Payroll</h1>
                <p className="subtitle">Gestión de Cobros - Tesorería</p>
              </div>
            </div>
            <div className="welcome-box">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/>
                <circle cx="12" cy="7" r="4"/>
              </svg>
              <span>Bienvenido, Usuario</span>
            </div>
          </div>
        </div>
      </header>

      {/* Contenido Principal */}
      <main className="main-content">
        <div className="container">
          
          {/* Sección de ayuda colapsable (opcional - puedes eliminarla si no la necesitas) */}
          <div className="help-section">
            <button 
              className="help-toggle"
              onClick={() => setIsHelpExpanded(!isHelpExpanded)}
              aria-expanded={isHelpExpanded}
            >
              <div className="help-toggle-header">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="12" cy="12" r="10"/>
                  <line x1="12" y1="16" x2="12" y2="12"/>
                  <line x1="12" y1="8" x2="12.01" y2="8"/>
                </svg>
                <span>¿Cómo usar esta aplicación?</span>
              </div>
              <svg 
                className={`chevron ${isHelpExpanded ? 'expanded' : ''}`}
                width="20" 
                height="20" 
                viewBox="0 0 24 24" 
                fill="none" 
                stroke="currentColor" 
                strokeWidth="2"
              >
                <polyline points="6 9 12 15 18 9"/>
              </svg>
            </button>
            <div className={`help-content ${isHelpExpanded ? 'expanded' : ''}`}>
              <ol className="help-list">
                <li>
                  <span className="step-number">1</span>
                  <div>
                    <strong>Paso 1</strong>
                    <p>Descripción del primer paso</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">2</span>
                  <div>
                    <strong>Paso 2</strong>
                    <p>Descripción del segundo paso</p>
                  </div>
                </li>
                <li>
                  <span className="step-number">3</span>
                  <div>
                    <strong>Paso 3</strong>
                    <p>Descripción del tercer paso</p>
                  </div>
                </li>
              </ol>
            </div>
          </div>

          {/* Card Principal - Aquí va tu contenido específico */}
          <div className="card">
            <div className="card-header">
              <h2>Gestión de Cuentas por Cobrar</h2>
              <p className="description">
                Sube el archivo Excel con la plantilla SIGO para gestionar los cobros por cliente.
              </p>
            </div>

            <div className="card-body">
              <div className="form-section">
                
                {/* Subida de archivo Excel - Cuentas por cobrar plantilla SIGO */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                      <polyline points="14 2 14 8 20 8"/>
                      <line x1="16" y1="13" x2="8" y2="13"/>
                      <line x1="16" y1="17" x2="8" y2="17"/>
                    </svg>
                    Cuentas por cobrar plantilla - SIGO
                  </label>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="file-input"
                    id="excel-upload"
                    onChange={(e) => setExcelFile(e.target.files[0] || null)}
                  />
                  <div
                    className={`drop-zone ${dragActive ? 'drag-active' : ''} ${excelFile ? 'has-file' : ''}`}
                    onClick={() => !excelFile && document.getElementById('excel-upload').click()}
                    onDragOver={(e) => { e.preventDefault(); setDragActive(true) }}
                    onDragLeave={() => setDragActive(false)}
                    onDrop={(e) => {
                      e.preventDefault()
                      setDragActive(false)
                      const file = e.dataTransfer.files[0]
                      if (file) setExcelFile(file)
                    }}
                  >
                    {excelFile ? (
                      <div className="file-preview">
                        <div className="file-icon">
                          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                            <polyline points="14 2 14 8 20 8"/>
                          </svg>
                        </div>
                        <div className="file-details">
                          <div className="file-name">{excelFile.name}</div>
                          <div className="file-size">{(excelFile.size / 1024).toFixed(1)} KB</div>
                        </div>
                        <button
                          className="btn-remove"
                          onClick={(e) => { e.stopPropagation(); setExcelFile(null) }}
                          title="Quitar archivo"
                        >
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      </div>
                    ) : (
                      <div className="drop-zone-content">
                        <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                          <polyline points="17 8 12 3 7 8"/>
                          <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <div className="drop-zone-text">
                          <span className="drop-zone-title">Arrastra tu archivo Excel aquí</span>
                          <span className="drop-zone-subtitle">o haz clic para seleccionarlo</span>
                        </div>
                        <span className="drop-zone-hint">.xlsx, .xls, .csv — Máximo 10 MB</span>
                      </div>
                    )}
                  </div>
                </div>

                {/* Multiselect de clientes */}
                <div className="form-group">
                  <label className="label">
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/>
                      <circle cx="9" cy="7" r="4"/>
                      <path d="M23 21v-2a4 4 0 0 0-3-3.87"/>
                      <path d="M16 3.13a4 4 0 0 1 0 7.75"/>
                    </svg>
                    Clientes
                  </label>
                  {clientes.length > 0 && (
                    <div className="multiselect-toolbar">
                      <button type="button" className="btn-link" onClick={() => setClientesSeleccionados(new Set(clientes))}>Seleccionar todos</button>
                      <span className="multiselect-sep">·</span>
                      <button type="button" className="btn-link" onClick={() => setClientesSeleccionados(new Set())}>Deseleccionar todos</button>
                      <span className="multiselect-count">{clientesSeleccionados.size} de {clientes.length} seleccionado{clientesSeleccionados.size !== 1 ? 's' : ''}</span>
                    </div>
                  )}
                  {clientes.length > 0 && (
                    <div className="search-clientes">
                      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <circle cx="11" cy="11" r="8"/>
                        <line x1="21" y1="21" x2="16.65" y2="16.65"/>
                      </svg>
                      <input
                        type="text"
                        placeholder="Buscar cliente..."
                        value={busquedaCliente}
                        onChange={(e) => setBusquedaCliente(e.target.value)}
                        className="search-clientes-input"
                      />
                      {busquedaCliente && (
                        <button type="button" className="search-clientes-clear" onClick={() => setBusquedaCliente('')}>
                          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <line x1="18" y1="6" x2="6" y2="18"/>
                            <line x1="6" y1="6" x2="18" y2="18"/>
                          </svg>
                        </button>
                      )}
                    </div>
                  )}
                  <div className={`checkbox-list ${!clientes.length ? 'checkbox-list--disabled' : ''}`}>
                    {clientes.length === 0 ? (
                      <span className="hint">Sube un archivo Excel para ver los clientes disponibles</span>
                    ) : (
                      clientes
                        .filter(c => c.toLowerCase().includes(busquedaCliente.toLowerCase()))
                        .map(cliente => (
                          <label key={cliente} className="checkbox-item">
                            <input
                              type="checkbox"
                              checked={clientesSeleccionados.has(cliente)}
                              onChange={(e) => {
                                const next = new Set(clientesSeleccionados)
                                e.target.checked ? next.add(cliente) : next.delete(cliente)
                                setClientesSeleccionados(next)
                              }}
                            />
                            <span>{cliente}</span>
                          </label>
                        ))
                    )}
                  </div>
                </div>

                {/* Botón Generar Formato */}
                <button
                  className="btn-primary"
                  onClick={generarFormato}
                  disabled={!excelFile || clientesSeleccionados.size === 0 || generando}
                >
                  {generando ? (
                    <>
                      <svg className="spinner" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
                      </svg>
                      Generando...
                    </>
                  ) : (
                    <>
                      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                        <path d="M12 18v-6M9 15l3 3 3-3"/>
                      </svg>
                      Generar Formato
                    </>
                  )}
                </button>

              </div>
            </div>
          </div>
        </div>
      </main>

      {/* Footer */}
      <footer className="footer">
        <div className="container">
          <p>&copy; {new Date().getFullYear()} Solutions & Payroll. Todos los derechos reservados.</p>
        </div>
      </footer>
    </div>
  )
}

export default App