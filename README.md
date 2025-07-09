# UBC Course Schedule to Calendar Converter

A web application that converts UBC course schedule Excel files to ICS format for easy import into Google Calendar or Apple Calendar.

## Features

- 🎯 **Simple Upload**: Drag and drop or click to upload Excel files
- 📅 **Calendar Ready**: Generates ICS files compatible with all major calendar apps
- 🎨 **UBC Branded**: Uses UBC colors and styling
- 📱 **Responsive**: Works on desktop and mobile devices
- ⚡ **Fast**: Real-time conversion with detailed error messages
- 🔒 **Secure**: Files are processed server-side and not stored

## Tech Stack

- **Frontend**: Next.js 14, React 18, TypeScript, Tailwind CSS
- **Backend**: Next.js API Routes
- **File Processing**: Python with pandas, openpyxl, ics library
- **Deployment**: Vercel

## Local Development

### Prerequisites

- Node.js 18+ 
- Python 3.8+
- npm or yarn

### Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd ubc-course-converter
   ```

2. **Install Node.js dependencies**
   ```bash
   npm install
   ```

3. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Create uploads directory**
   ```bash
   mkdir uploads
   ```

5. **Run the development server**
   ```bash
   npm run dev
   ```

6. **Open your browser**
   Navigate to [http://localhost:3000](http://localhost:3000)

## Deployment to Vercel

### Automatic Deployment

1. **Connect to Vercel**
   - Push your code to GitHub
   - Connect your repository to Vercel
   - Vercel will automatically detect Next.js and deploy

2. **Environment Setup**
   - Vercel will automatically install Node.js dependencies
   - Python dependencies are handled via `requirements.txt`

### Manual Deployment

1. **Install Vercel CLI**
   ```bash
   npm i -g vercel
   ```

2. **Deploy**
   ```bash
   vercel
   ```

3. **Follow the prompts**
   - Link to existing project or create new
   - Set up custom domain (optional)

## How It Works

### For Users

1. **Export Course Schedule**: From UBC SSC, export your course schedule as Excel
2. **Upload File**: Drag and drop or click to upload the Excel file
3. **Download ICS**: The converter generates an ICS file for download
4. **Import to Calendar**: Import the ICS file into Google Calendar or Apple Calendar

### Technical Process

1. **File Upload**: Frontend sends Excel file to Next.js API route
2. **File Processing**: API saves file and calls Python script
3. **Excel Parsing**: Python script reads Excel file using pandas
4. **Data Extraction**: Extracts course info, times, locations, instructors
5. **ICS Generation**: Creates ICS calendar file with recurring events
6. **File Download**: Returns ICS file to user for download

## File Structure

```
ubc-course-converter/
├── app/                          # Next.js app directory
│   ├── api/                      # API routes
│   │   └── convert/              # File conversion endpoint
│   ├── globals.css               # Global styles
│   ├── layout.tsx                # Root layout
│   └── page.tsx                  # Main page component
├── excel_to_ics_web.py          # Python conversion script
├── requirements.txt              # Python dependencies
├── package.json                  # Node.js dependencies
├── tailwind.config.js           # Tailwind configuration
├── vercel.json                   # Vercel configuration
└── README.md                     # This file
```

## Error Handling

The application provides detailed error messages for common issues:

- **Invalid File Format**: Ensures only Excel files are uploaded
- **Missing Data**: Checks for required UBC course schedule columns
- **Processing Errors**: Handles Python script execution issues
- **No Events Found**: Validates that course data was successfully extracted

## Customization

### Styling
- Modify `tailwind.config.js` for color changes
- Update `app/globals.css` for custom styles
- UBC colors are defined in the Tailwind config

### Error Messages
- Edit error handling in `app/api/convert/route.ts`
- Update user-facing messages in `app/page.tsx`

### Python Script
- Modify `excel_to_ics_web.py` for different Excel formats
- Add support for other universities by updating column names

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is open source and available under the MIT License.

## Support

For issues or questions:
- Create an issue on GitHub
- Check the troubleshooting section on the website
- Ensure you're using a valid UBC course schedule export

---

Built for UBC students by UBC students 🎓 