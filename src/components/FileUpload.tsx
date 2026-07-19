'use client'

import { useCallback, useRef, useState } from 'react'
import { Upload, FileSpreadsheet, X, Loader2 } from 'lucide-react'
import { cn } from '@/lib/utils'

interface FileUploadProps {
  onFileParsed: (result: { fileName: string; textRepresentation: string } | null) => void
  disabled?: boolean
}

export function FileUpload({ onFileParsed, disabled }: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [file, setFile] = useState<File | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [loading, setLoading] = useState(false)

  const handleFile = useCallback(async (f: File) => {
    setError(null)
    const ext = f.name.split('.').pop()?.toLowerCase()
    if (!ext || !['xlsx', 'xls', 'csv'].includes(ext)) {
      setError('Format accepté : .xlsx, .xls, .csv')
      return
    }
    if (f.size > 10 * 1024 * 1024) {
      setError('Fichier trop volumineux (max 10 Mo)')
      return
    }
    setLoading(true)
    try {
      const { parseFile } = await import('@/lib/fileParser')
      const result = await parseFile(f)
      setFile(f)
      onFileParsed({
        fileName: f.name,
        textRepresentation: result.textRepresentation,
      })
    } catch (err) {
      console.error(err)
      setError('Erreur lors de la lecture du fichier')
      onFileParsed(null)
    } finally {
      setLoading(false)
    }
  }, [onFileParsed])

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault()
      if (disabled || loading) return
      const f = e.dataTransfer.files[0]
      if (f) handleFile(f)
    },
    [handleFile, disabled, loading]
  )

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      if (disabled || loading) return
      const f = e.target.files?.[0]
      if (f) handleFile(f)
    },
    [handleFile, disabled, loading]
  )

  const remove = useCallback((e: React.MouseEvent) => {
    e.stopPropagation()
    setFile(null)
    setError(null)
    onFileParsed(null)
    if (inputRef.current) inputRef.current.value = ''
  }, [onFileParsed])

  if (file) {
    return (
      <div className="flex items-center gap-2 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-2 text-sm dark:border-emerald-800/50 dark:bg-emerald-950/20">
        <FileSpreadsheet className="h-4 w-4 text-emerald-600 dark:text-emerald-400" />
        <span className="flex-1 truncate text-emerald-700 dark:text-emerald-300">
          {file.name}
        </span>
        <button
          onClick={remove}
          type="button"
          className="rounded p-1 text-emerald-500 hover:bg-emerald-100 dark:hover:bg-emerald-900/30 transition-colors"
          title="Retirer le fichier"
        >
          <X className="h-3.5 w-3.5" />
        </button>
      </div>
    )
  }

  return (
    <div className="w-full">
      <div
        onDrop={handleDrop}
        onDragOver={(e) => e.preventDefault()}
        onClick={() => !disabled && !loading && inputRef.current?.click()}
        className={cn(
          'flex cursor-pointer items-center justify-center gap-2 rounded-xl border border-dashed px-4 py-3 text-sm transition-all duration-200',
          error
            ? 'border-destructive/40 bg-destructive/5 text-destructive dark:border-destructive/30'
            : 'border-border/60 text-muted-foreground hover:border-primary/40 hover:bg-muted/30 hover:text-foreground',
          (disabled || loading) && 'pointer-events-none opacity-50'
        )}
      >
        {loading ? (
          <>
            <Loader2 className="h-4 w-4 animate-spin text-primary" />
            <span className="text-xs">Lecture du fichier...</span>
          </>
        ) : (
          <>
            <Upload className="h-4 w-4 shrink-0 text-muted-foreground" />
            <span className="truncate">{error || 'Glisser ou cliquer pour ajouter un fichier Excel (.xlsx, .csv)'}</span>
          </>
        )}
      </div>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,.csv"
        className="hidden"
        onChange={handleChange}
        disabled={disabled || loading}
      />
    </div>
  )
}
