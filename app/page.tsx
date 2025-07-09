'use client'

import { useState, useCallback } from 'react'
import axios from 'axios'

interface ConversionResult {
  success: boolean
  message: string
  downloadUrl?: string
}

export default function Home() {
  const [isDragOver, setIsDragOver] = useState(false)
  const [isConverting, setIsConverting] = useState(false)
  const [result, setResult] = useState<ConversionResult | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [skipBreaks, setSkipBreaks] = useState(false)

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(true)
  }, [])

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(false)
  }, [])

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    setIsDragOver(false)
    
    const files = Array.from(e.dataTransfer.files)
    const excelFile = files.find(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    )
    
    if (excelFile) {
      handleFileUpload(excelFile)
    } else {
      setError('Please upload an Excel file (.xlsx or .xls)')
    }
  }, [skipBreaks])

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0]
    if (file) {
      handleFileUpload(file)
    }
  }, [skipBreaks])

  const handleFileUpload = async (file: File) => {
    setIsConverting(true)
    setError(null)
    setResult(null)

    const formData = new FormData()
    formData.append('file', file)
    formData.append('skip_breaks', skipBreaks ? '1' : '0')

    try {
      const response = await axios.post('/api/convert', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob',
      })

      // Create download link
      const blob = new Blob([response.data], { type: 'text/calendar' })
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'courses.ics'
      document.body.appendChild(a)
      a.click()
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)

      setResult({
        success: true,
        message: 'Conversion successful! Your ICS file has been downloaded.',
        downloadUrl: url
      })
    } catch (err: any) {
      let errorMessage = 'An error occurred during conversion.'
      
      if (err.response?.data) {
        const reader = new FileReader()
        reader.onload = () => {
          try {
            const errorData = JSON.parse(reader.result as string)
            errorMessage = errorData.error || errorMessage
          } catch {
            errorMessage = 'Invalid file format or structure.'
          }
          setError(errorMessage)
        }
        reader.readAsText(err.response.data)
      } else {
        setError(errorMessage)
      }
    } finally {
      setIsConverting(false)
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-ubc-blue to-blue-900 text-white">
      <div className="container mx-auto px-4 py-8">
        <div className="max-w-4xl mx-auto">
          {/* Header */}
          <div className="text-center mb-8">
            <h1 className="text-4xl font-bold mb-4">
              UBC Course Schedule to Calendar Converter
            </h1>
          </div>

          {/* Beta Toggle */}
          <div className="flex items-center justify-center mb-4">
            <label className="flex items-center cursor-pointer gap-2">
              <input
                type="checkbox"
                checked={skipBreaks}
                onChange={e => setSkipBreaks(e.target.checked)}
                className="form-checkbox h-5 w-5 text-ubc-gold"
                disabled={isConverting}
              />
              <span className="font-medium">Skip UBC Holidays/Reading Breaks</span>
              <span className="ml-2 px-2 py-0.5 text-xs rounded bg-yellow-400 text-yellow-900 font-bold uppercase">Beta</span>
            </label>
          </div>

          {/* Upload Area */}
          <div className="mb-8">
            <div
              className={`relative border-2 border-dashed rounded-lg p-8 text-center transition-colors ${
                isDragOver 
                  ? 'border-ubc-gold bg-ubc-gold/10' 
                  : 'border-gray-400 hover:border-ubc-gold'
              }`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileSelect}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                disabled={isConverting}
              />
              
              <div className="space-y-4">
                <div className="text-6xl mb-4">📅</div>
                <h3 className="text-2xl font-semibold">
                  {isConverting ? 'Converting...' : 'Upload Your UBC Course Schedule'}
                </h3>
                <p className="text-gray-300">
                  Drag and drop your Excel file here, or click to browse
                </p>
                <p className="text-sm text-gray-400">
                  Only .xlsx and .xls files are supported
                </p>
              </div>
            </div>
          </div>

          {/* Error Message */}
          {error && (
            <div className="mb-6 p-4 bg-red-500/20 border border-red-500 rounded-lg">
              <h4 className="font-semibold text-red-300 mb-2">Error</h4>
              <p className="text-red-200">{error}</p>
            </div>
          )}

          {/* Success Message */}
          {result?.success && (
            <div className="mb-6 p-4 bg-green-500/20 border border-green-500 rounded-lg">
              <h4 className="font-semibold text-green-300 mb-2">Success!</h4>
              <p className="text-green-200">{result.message}</p>
            </div>
          )}

          {/* Instructions */}
          <div className="bg-white/10 backdrop-blur-sm rounded-lg p-6">
            <h2 className="text-2xl font-bold mb-4">How to Use</h2>
            
            <div className="space-y-6">
              <div>
                <h3 className="text-lg font-semibold mb-2">1. Export Your Course Schedule</h3>
                <p className="text-gray-200">
                  From UBC Workday, go to Academics &gt; Registration & Courses &gt; View My Courses, then click the "Export to Excel" button.
                </p>
              </div>

              <div>
                <h3 className="text-lg font-semibold mb-2">2. Upload and Convert</h3>
                <p className="text-gray-200">
                  Upload your Excel file using the tool above. The converter will extract your course information and create an ICS file.
                </p>
              </div>

              <div>
                <h3 className="text-lg font-semibold mb-2">3. Import to Your Calendar</h3>
                
                <div className="space-y-4">
                  <div>
                    <h4 className="font-medium text-ubc-gold">Google Calendar:</h4>
                    <ol className="list-decimal list-inside text-gray-200 space-y-1 ml-4">
                      <li>Open Google Calendar</li>
                      <li>Go to Settings</li>
                      <li>Go to "Import & export"</li>
                      <li>Choose the downloaded ICS file</li>
                      <li>Select your preferred calendar</li>
                      <li>Click "Import"</li>
                    </ol>
                  </div>

                  
                </div>
              </div>
            </div>
          </div>

          
        </div>
      </div>
    </div>
  )
} 