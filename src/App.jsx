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
      ws.properties.outlineProperties = { summaryBelow: true, summaryRight: false }

      // Mapear encabezados de fila 1 en la plantilla (+ "Total cartera" para el subtotal)
      const colMap = {}
      ws.getRow(1).eachCell({ includeEmpty: false }, (cell, col) => {
        const val = String(cell.value ?? '').trim()
        if (CAMPOS.includes(val) || val === 'Total cartera' || val === '%') colMap[val] = col
      })

      // Capturar estilos de referencia desde fila 2 (si existe) para aplicar a nuevas filas
      const refStyles = {}
      for (const campo of [...CAMPOS, 'Total cartera', '%']) {
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

      // Letra de columna "Total cartera" para la fórmula SUBTOTAL (1-based en ExcelJS → 0-based en xlsx)
      const colTotalCartera = colMap['Total cartera']
      const colTotalCarteraLetra = colTotalCartera ? XLSX.utils.encode_col(colTotalCartera - 1) : null
      const colPrimerVencidoLetra = colMap['Vencido 1 a 30'] ? XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1) : null
      const colSaldoVencerLetra = colMap['Saldo por vencer'] ? XLSX.utils.encode_col(colMap['Saldo por vencer'] - 1) : null

      // Agrupar dataRows por cliente (grupos consecutivos)
      const grupos = []
      for (const row of dataRows) {
        const cliente = String(row['Cliente'] ?? '').trim()
        const ultimo = grupos[grupos.length - 1]
        if (ultimo && ultimo.cliente === cliente) {
          ultimo.rows.push(row)
        } else {
          grupos.push({ cliente, rows: [row] })
        }
      }

      // Escribir por grupo: filas de datos + fila de subtotal por cliente
      let currentExcelRow = 2
      const subtotalHLRows = []  // filas de subtotal cuyo cliente contiene "HL"
      const subtotalAllRows = []  // todas las filas de subtotal (para fórmula %)
      const clienteSubtotalRow = {}  // cliente (normalizado) → número de fila de subtotal
      for (const { cliente, rows } of grupos) {
        const startRow = currentExcelRow

        for (const rowData of rows) {
          const excelRow = ws.getRow(currentExcelRow)
          for (const campo of CAMPOS) {
            if (!colMap[campo]) continue
            const cell = excelRow.getCell(colMap[campo])
            if (refStyles[colMap[campo]]) cell.style = JSON.parse(JSON.stringify(refStyles[colMap[campo]]))
            cell.value = rowData[campo] ?? null
          }
          // Fórmula SUMA en columna Total cartera
          if (colTotalCartera && colPrimerVencidoLetra && colSaldoVencerLetra) {
            const cell = excelRow.getCell(colTotalCartera)
            if (refStyles[colTotalCartera]) cell.style = JSON.parse(JSON.stringify(refStyles[colTotalCartera]))
            cell.value = { formula: `SUM(${colPrimerVencidoLetra}${currentExcelRow}:${colSaldoVencerLetra}${currentExcelRow})` }
          }
          excelRow.outlineLevel = 1
          excelRow.commit()
          currentExcelRow++
        }

        // Fila de subtotal
        const endRow = currentExcelRow - 1
        const subtotalRow = ws.getRow(currentExcelRow)
        if (colMap['Cliente']) {
          const cell = subtotalRow.getCell(colMap['Cliente'])
          if (refStyles[colMap['Cliente']]) cell.style = JSON.parse(JSON.stringify(refStyles[colMap['Cliente']]))
          cell.font = { ...(cell.font || {}), bold: true }
          cell.value = `Total ${cliente}`
        }
        for (const campo of ['Total cartera', 'Vencido 1 a 30', 'Vencido 31 a 60', 'Vencido 61 a 90', 'Vencido más de 91', 'Saldo por vencer']) {
          if (!colMap[campo]) continue
          const colIdx = colMap[campo]
          const colLetra = XLSX.utils.encode_col(colIdx - 1)
          const cell = subtotalRow.getCell(colIdx)
          if (refStyles[colIdx]) cell.style = JSON.parse(JSON.stringify(refStyles[colIdx]))
          cell.font = { ...(cell.font || {}), bold: true }
          cell.value = { formula: `SUBTOTAL(9,${colLetra}${startRow}:${colLetra}${endRow})` }
        }
        subtotalRow.commit()
        subtotalAllRows.push(currentExcelRow)
        clienteSubtotalRow[cliente.trim().toUpperCase()] = currentExcelRow
        if (/HL/.test(cliente)) subtotalHLRows.push(currentExcelRow)
        currentExcelRow++
      }

      // Escribir fórmula % en todas las filas de subtotal (necesita firstFijaRow → se hace aquí)
      if (colMap['%'] && colMap['Total cartera']) {
        const colPct = colMap['%']
        const colTCLetra = XLSX.utils.encode_col(colMap['Total cartera'] - 1)
        const fijaRow7 = currentExcelRow + 1  // firstFijaRow calculado igual que abajo
        const refTC = `$${colTCLetra}$${fijaRow7}`
        for (const rowNum of subtotalAllRows) {
          const tcLetraRow = `${colTCLetra}${rowNum}`
          const pctCell = ws.getRow(rowNum).getCell(colPct)
          if (refStyles[colPct]) pctCell.style = JSON.parse(JSON.stringify(refStyles[colPct]))
          pctCell.font = { ...(pctCell.font || {}), bold: true }
          pctCell.numFmt = '0.00%'
          pctCell.value = { formula: `IF(${refTC}=0,0,${tcLetraRow}/${refTC})` }
          ws.getRow(rowNum).commit()
        }
      }

      // Agregar fila vacía de separación y luego las filas fijas al final
      const lastDataRow = currentExcelRow - 1  // última fila con datos (último "Total [cliente]")
      const firstFijaRow = currentExcelRow + 1
      const separador = ws.getRow(currentExcelRow)
      separador.commit()

      // Columnas que llevan SUBTOTAL dinámico en la primera fila fija
      const camposSubtotal = ['Total cartera', 'Vencido 1 a 30', 'Vencido 31 a 60', 'Vencido 61 a 90', 'Vencido más de 91', 'Saldo por vencer']
      const subtotalCols = {}
      for (const campo of camposSubtotal) {
        if (colMap[campo]) {
          subtotalCols[colMap[campo]] = XLSX.utils.encode_col(colMap[campo] - 1)
        }
      }

      const mesesES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre']
      const mesActual = mesesES[new Date().getMonth()]

      fijasCaptured.forEach((fija, i) => {
        const excelRow = ws.getRow(firstFijaRow + i)
        if (fija.height) excelRow.height = fija.height
        fija.cells.forEach(({ col, value, style, numFmt }) => {
          const cell = excelRow.getCell(col)
          // En la primera fila fija (fila 7): SUBTOTAL de todo el rango de datos
          if (i === 0 && subtotalCols[col]) {
            cell.value = { formula: `SUBTOTAL(9,${subtotalCols[col]}2:${subtotalCols[col]}${lastDataRow})` }
          // En la primera fila fija (fila 7), columna "%": suma del rango completo
          } else if (i === 0 && colMap['%'] && col === colMap['%']) {
            const colPctLetra = XLSX.utils.encode_col(colMap['%'] - 1)
            cell.value = { formula: `SUM(${colPctLetra}2:${colPctLetra}${lastDataRow})` }
            cell.numFmt = '0.00%'
          // En la segunda fila fija (fila 8): SUM solo de las filas "Total HL..."
          } else if (i === 1 && subtotalCols[col] && subtotalHLRows.length > 0) {
            const refs = subtotalHLRows.map(r => `${subtotalCols[col]}${r}`).join(',')
            cell.value = { formula: `SUM(${refs})` }
          // En la segunda fila fija (fila 8), columna Cliente: mes dinámico
          } else if (i === 1 && colMap['Cliente'] && col === colMap['Cliente']) {
            cell.value = `Cartera Grupo HL ${mesActual}`
          // En la tercera fila fija (fila 9), columna "Vencido 1 a 30": suma de Vencido 31/61/91 de la fila 7
          } else if (i === 2 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const c31 = subtotalCols[colMap['Vencido 31 a 60']]
            const c61 = subtotalCols[colMap['Vencido 61 a 90']]
            const c91 = subtotalCols[colMap['Vencido más de 91']]
            if (c31 && c61 && c91) {
              cell.value = { formula: `SUM(${c31}${firstFijaRow},${c61}${firstFijaRow},${c91}${firstFijaRow})` }
            } else {
              cell.value = value
            }
          // En la cuarta fila fija (fila 10), columna "Vencido 1 a 30": suma de Vencido 31/61/91 de la fila 8
          } else if (i === 3 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const c31 = subtotalCols[colMap['Vencido 31 a 60']]
            const c61 = subtotalCols[colMap['Vencido 61 a 90']]
            const c91 = subtotalCols[colMap['Vencido más de 91']]
            if (c31 && c61 && c91) {
              cell.value = { formula: `SUM(${c31}${firstFijaRow + 1},${c61}${firstFijaRow + 1},${c91}${firstFijaRow + 1})` }
            } else {
              cell.value = value
            }
          // En la cuarta fila fija (fila 10), columna "%": =IF($G$fila9=0,0,G_fila10/$G$fila9)
          } else if (i === 3 && colMap['%'] && col === colMap['%']) {
            const colV130Letra = XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1)
            const fila9 = firstFijaRow + 2
            const fila10 = firstFijaRow + 3
            cell.value = { formula: `IF($${colV130Letra}$${fila9}=0,0,${colV130Letra}${fila10}/$${colV130Letra}$${fila9})` }
            cell.numFmt = '0.00%'
          // Fila 13 (i=6), columna "Vencido 1 a 30": fila9 - fila10
          } else if (i === 6 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const colLetra = XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1)
            const fila9 = firstFijaRow + 2
            const fila10 = firstFijaRow + 3
            cell.value = { formula: `${colLetra}${fila9}-${colLetra}${fila10}` }
          // Fila 18 (i=11), columna "Vencido 1 a 30": fila13 - SUMA(fila14:fila17)
          } else if (i === 11 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const colLetra = XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1)
            const fila13 = firstFijaRow + 6
            const fila14 = firstFijaRow + 7
            const fila17 = firstFijaRow + 10
            cell.value = { formula: `${colLetra}${fila13}-SUM(${colLetra}${fila14}:${colLetra}${fila17})` }
          // Filas 14-17 (i=7..10), columna "Vencido 1 a 30": suma Vencido31+61+91 del subtotal del cliente específico
          } else if (i >= 7 && i <= 10 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const clientesFijas = [
              'PUNTO MEDICAL DISTRIBUCIONES SAS',
              'REJIMETAL SAS',
              'DINTERWEB SAS',
              'SOLUTIONS & PAYROLL PERU S.A.C',
            ]
            const clienteTarget = clientesFijas[i - 7]
            const subtotalFila = clienteSubtotalRow[clienteTarget.toUpperCase()]
            const c31 = subtotalCols[colMap['Vencido 31 a 60']]
            const c61 = subtotalCols[colMap['Vencido 61 a 90']]
            const c91 = subtotalCols[colMap['Vencido más de 91']]
            if (subtotalFila && c31 && c61 && c91) {
              cell.value = { formula: `SUM(${c31}${subtotalFila},${c61}${subtotalFila},${c91}${subtotalFila})` }
            } else {
              cell.value = value
            }
          // Filas 14-18 (i=7..11), columna "%": =IF($G$fila13=0,0,G{filaActual}/$G$fila13)
          } else if (i >= 7 && i <= 11 && colMap['%'] && col === colMap['%']) {
            const colV130Letra = XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1)
            const fila13 = firstFijaRow + 6
            const filaActual = firstFijaRow + i
            cell.value = { formula: `IF($${colV130Letra}$${fila13}=0,0,${colV130Letra}${filaActual}/$${colV130Letra}$${fila13})` }
            cell.numFmt = '0.00%'
          // Fila 21 (i=14), columna "Vencido 1 a 30": fila10 + fila18
          } else if (i === 14 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const colLetra = XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1)
            cell.value = { formula: `${colLetra}${firstFijaRow + 3}+${colLetra}${firstFijaRow + 11}` }
          // Fila 22 (i=15), columna "Vencido 1 a 30": SUMA(fila14:fila17)
          } else if (i === 15 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const colLetra = XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1)
            cell.value = { formula: `SUM(${colLetra}${firstFijaRow + 7}:${colLetra}${firstFijaRow + 10})` }
          // Fila 23 (i=16), columna "Vencido 1 a 30": SUMA(fila21:fila22)
          } else if (i === 16 && colMap['Vencido 1 a 30'] && col === colMap['Vencido 1 a 30']) {
            const colLetra = XLSX.utils.encode_col(colMap['Vencido 1 a 30'] - 1)
            cell.value = { formula: `SUM(${colLetra}${firstFijaRow + 14}:${colLetra}${firstFijaRow + 15})` }
          } else {
            cell.value = value
          }
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