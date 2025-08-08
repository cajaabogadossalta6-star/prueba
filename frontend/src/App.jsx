import { useState, useEffect } from 'react'
import './App.css'

function App() {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [parametros, setParametros] = useState('')
  const [taskId, setTaskId] = useState(null)
  const [progress, setProgress] = useState(null)
  const [status, setStatus] = useState('')
  const [pdfUrl, setPdfUrl] = useState('')
  const [excelUrl, setExcelUrl] = useState('')
  const [error, setError] = useState('')

  const iniciarDescarga = async (e) => {
    e.preventDefault()
    setError('')
    setProgress(null)
    setStatus('')
    setPdfUrl('')
    setExcelUrl('')
    try {
      const resp = await fetch('/descargas', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ usuario: username, clave: password, parametros })
      })
      if (!resp.ok) throw new Error('Error iniciando descarga')
      const data = await resp.json()
      setTaskId(data.id)
    } catch (err) {
      setError(err.message)
    }
  }

  useEffect(() => {
    if (!taskId) return
    const interval = setInterval(async () => {
      try {
        const resp = await fetch(`/tareas/${taskId}`)
        if (!resp.ok) throw new Error('Error consultando tarea')
        const data = await resp.json()
        setProgress(data.progreso ?? data.progress ?? 0)
        setStatus(data.estado ?? data.status)
        if (data.pdf) setPdfUrl(data.pdf)
        if (data.excel) setExcelUrl(data.excel)
        if (data.status === 'completado' || data.status === 'done') {
          clearInterval(interval)
        }
      } catch (err) {
        setError(err.message)
        clearInterval(interval)
      }
    }, 2000)
    return () => clearInterval(interval)
  }, [taskId])

  return (
    <div className="container">
      <h1>Descargas</h1>
      <form className="form" onSubmit={iniciarDescarga}>
        <input
          type="text"
          placeholder="Usuario"
          value={username}
          onChange={e => setUsername(e.target.value)}
          required
        />
        <input
          type="password"
          placeholder="Contraseña"
          value={password}
          onChange={e => setPassword(e.target.value)}
          required
        />
        <input
          type="text"
          placeholder="Parámetros de descarga"
          value={parametros}
          onChange={e => setParametros(e.target.value)}
        />
        <button type="submit">Iniciar descarga</button>
      </form>
      {error && <p className="error">{error}</p>}
      {status && <p>Estado: {status}</p>}
      {progress !== null && (
        <div className="progress">
          <div className="bar" style={{ width: `${progress}%` }} />
          <span>{progress}%</span>
        </div>
      )}
      {pdfUrl && <p><a href={pdfUrl}>Descargar PDF</a></p>}
      {excelUrl && <p><a href={excelUrl}>Descargar Excel</a></p>}
    </div>
  )
}

export default App
