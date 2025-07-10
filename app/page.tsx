'use client'

import { useState, useCallback } from 'react'
import { useDropzone } from 'react-dropzone'

export default function Home() {
  const [file, setFile] = useState<File | null>(null)
  const [isConverting, setIsConverting] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [success, setSuccess] = useState(false)

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop: useCallback((acceptedFiles: File[]) => {
      const selectedFile = acceptedFiles[0]
      setFile(selectedFile)
      setError(null)
    }, []),
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
    },
    multiple: false,
  })

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    
    if (!file) {
      setError('Please select a file to upload')
      return
    }

    setIsConverting(true)
    setError(null)
    
    try {
      const formData = new FormData()
      formData.append('file', file)
      
      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
      })
      
      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Failed to convert file')
      }
      
      // Get the blob data
      const blob = await response.blob()
      
      // Create a download link
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.style.display = 'none'
      a.href = url
      a.download = 'courses.ics'
      document.body.appendChild(a)
      a.click()
      
      // Clean up
      window.URL.revokeObjectURL(url)
      document.body.removeChild(a)
      
      setSuccess(true)
      setTimeout(() => setSuccess(false), 5000)
    } catch (err: any) {
      setError(err.message || 'An error occurred during conversion')
    } finally {
      setIsConverting(false)
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-800 to-blue-900 text-white">
      <div className="container mx-auto px-4 py-8">
        <div className="max-w-4xl mx-auto">
          {/* Header */}
          <div className="text-center mb-8">
            <h1 className="text-4xl font-bold mb-4">
              UBC Course Schedule to Calendar Converter
            </h1>
          </div>

          {/* Upload Area */}
          <div className="mb-8">
            <form onSubmit={handleSubmit}>
              <div 
                {...getRootProps()} 
                className={`relative border-2 border-dashed rounded-lg p-8 text-center transition-colors ${
                  isDragActive 
                    ? 'border-yellow-400 bg-yellow-400/10' 
                    : 'border-gray-400 hover:border-yellow-400'
                }`}
              >
                <input {...getInputProps()} />
                
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
              
              {file && (
                <div className="mt-4 p-3 bg-white/10 rounded flex items-center justify-between">
                  <span className="truncate max-w-xs">{file.name}</span>
                  <span className="text-sm text-gray-300">
                    {(file.size / 1024).toFixed(1)} KB
                  </span>
                </div>
              )}
              
              {file && (
                <button 
                  type="submit" 
                  disabled={isConverting}
                  className="mt-4 w-full py-2 px-4 rounded font-medium bg-yellow-400 hover:bg-yellow-500 text-blue-900 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  {isConverting ? 'Converting...' : 'Convert to Calendar'}
                </button>
              )}
            </form>
          </div>
          
          {/* Error Message */}
          {error && (
            <div className="mb-6 p-4 bg-red-500/20 border border-red-500 rounded-lg">
              <h4 className="font-semibold text-red-300 mb-2">Error</h4>
              <p className="text-red-200">{error}</p>
            </div>
          )}
          
          {/* Success Message */}
          {success && (
            <div className="mb-6 p-4 bg-green-500/20 border border-green-500 rounded-lg">
              <h4 className="font-semibold text-green-300 mb-2">Success!</h4>
              <p className="text-green-200">Conversion successful! Your calendar file has been downloaded.</p>
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
                    <h4 className="font-medium text-yellow-400">Google Calendar:</h4>
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