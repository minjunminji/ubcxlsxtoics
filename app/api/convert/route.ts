import { NextRequest, NextResponse } from 'next/server'
import { writeFile, mkdir } from 'fs/promises'
import { join } from 'path'
import { exec } from 'child_process'
import { promisify } from 'util'
import { existsSync } from 'fs'
import os from 'os'

const execAsync = promisify(exec)

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData()
    const file = formData.get('file') as File | null

    if (!file) {
      return NextResponse.json(
        { error: 'No file uploaded' },
        { status: 400 }
      )
    }

    // Validate extension
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
      return NextResponse.json(
        { error: 'Please upload an Excel file (.xlsx or .xls)' },
        { status: 400 }
      )
    }

    // ------------------ Save uploaded file ------------------
    const bytes = await file.arrayBuffer()
    const buffer = Buffer.from(bytes)

    // Use the writable /tmp directory provided by the serverless runtime
    const uploadDir = join(os.tmpdir(), 'uploads')
    if (!existsSync(uploadDir)) {
      // Recursively create /tmp/uploads if it does not exist
      await mkdir(uploadDir, { recursive: true })
    }

    const filePath = join(uploadDir, file.name)
    await writeFile(filePath, buffer)

    // ------------------ Invoke the Python converter ------------------
    const pythonScript = join(process.cwd(), 'app', 'api', 'convert', 'index.py')

    // Prefer python3 but fall back to python if python3 is unavailable
    let stdout: string, stderr: string | undefined
    try {
      ;({ stdout, stderr } = await execAsync(`python3 "${pythonScript}" "${filePath}"`))
    } catch (err: any) {
      // Retry with `python` if python3 failed (e.g., command not found)
      ;({ stdout, stderr } = await execAsync(`python "${pythonScript}" "${filePath}"`))
    }

    if (stderr) {
      console.error('Python script error:', stderr)
    }

    // If the python script returned JSON (error payload) treat it as an error
    if (stdout.trim().startsWith('{')) {
      try {
        const errorData = JSON.parse(stdout)
        return NextResponse.json(
          { error: errorData.error || 'Conversion failed' },
          { status: 400 }
        )
      } catch {
        // Not valid JSON, continue assuming ICS content
      }
    }

    // Success – return the ICS content
    return new NextResponse(stdout, {
      headers: {
        'Content-Type': 'text/calendar',
        'Content-Disposition': 'attachment; filename="courses.ics"',
      },
    })
  } catch (error: any) {
    console.error('Conversion error:', error)

    // Provide more user-friendly messages for common failures
    let errorMessage = 'An unknown error occurred on the server.'
    const msg = error.message || ''

    if (msg.includes('No module named')) {
      errorMessage = 'Server configuration error. Please try again later.'
    } else if (msg.includes('FileNotFoundError')) {
      errorMessage = 'Invalid file format. Please ensure you uploaded a valid UBC course schedule Excel file.'
    } else if (msg.includes('Permission denied')) {
      errorMessage = 'File access error. Please try again.'
    } else if (msg.includes('ENOENT')) {
      errorMessage = 'File system error. Please try again.'
    }

    return NextResponse.json(
      { error: errorMessage },
      { status: 500 }
    )
  }
} 