import fs from 'fs'
import path from 'path'
import ExcelJS from 'exceljs'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import { execSync } from 'child_process'

const templatePath = 'input/template.docx'
const outputDir = 'output'
const excelPath = 'input/companyList.xlsx'

// Ensure output directory exists
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir)

async function generateDocs() {
  try {
    // Read Excel file
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(excelPath)
    const worksheet = workbook.getWorksheet(1)

    // Read template
    const templateContent = fs.readFileSync(templatePath, 'binary')

    testDoc.render()
    console.log('✅ Template is valid')

    let successCount = 0
    let errorCount = 0

    for (const row of worksheet.getRows(2, worksheet.rowCount - 1) || []) {
      const companyNumber = row.getCell('B').text?.trim() || ''
      const companyInitial = row.getCell('C').text?.trim() || ''
      const companyHeader = row.getCell('A').text?.trim() || ''

      if (!companyNumber || !companyInitial || !companyHeader) {
        console.log(`⚠️  Skipping row with missing data: ${companyNumber}`)
        continue
      }

      try {
        // Create new document instance for each iteration
        const docZip = new PizZip(templateContent)
        const doc = new Docxtemplater(docZip, {
          paragraphLoop: true,
          linebreaks: true,
        })

        doc.setData({
          CompanyNumber: companyNumber,
          CompanyInitial: companyInitial,
          CompanyHeader: companyHeader,
        })

        doc.render()

        // Generate DOCX
        const docBuffer = doc.getZip().generate({ type: 'nodebuffer' })
        const docxFilename = `${companyNumber}_${companyInitial} Schedule 9_CC.docx`
        const docxPath = path.join(outputDir, docxFilename)
        fs.writeFileSync(docxPath, docBuffer)

        // Convert to PDF using LibreOffice CLI
        try {
          execSync(
            `soffice --headless --convert-to pdf --outdir "${outputDir}" "${docxPath}"`,
            { timeout: 30000 } // 30 second timeout
          )
          console.log(`✅ Generated: ${docxFilename}`)
          successCount++
        } catch (pdfError) {
          console.error(
            `❌ PDF conversion failed for ${docxFilename}:`,
            pdfError.message
          )
          console.log(`📄 DOCX file created successfully: ${docxFilename}`)
          successCount++
        }
      } catch (renderError) {
        console.error(
          `❌ Failed to process row ${companyNumber}:`,
          renderError.message
        )
        errorCount++
      }
    }

    console.log(`\n🎉 Process completed!`)
    console.log(`✅ Successful: ${successCount}`)
    console.log(`❌ Errors: ${errorCount}`)
  } catch (error) {
    if (error.name === 'TemplateError') {
      console.error('\n❌ TEMPLATE FORMATTING ERROR')
      console.error(
        'Your Word document has formatting issues that break the template variables.'
      )
      console.error('\n🔧 TO FIX:')
      console.error('1. Open mailmerge.docx in Microsoft Word')
      console.error('2. Select all text (Ctrl+A)')
      console.error('3. Remove formatting (Ctrl+Shift+N)')
      console.error(
        '4. Make sure variables like {{CompanyHeader}} are typed fresh'
      )
      console.error('5. Save and try again')

      if (error.properties?.errors) {
        console.error('\n📝 Specific issues:')
        error.properties.errors.forEach((err, i) => {
          console.error(`${i + 1}. ${err.message}`)
          console.error(
            `   Problem with: "${err.properties?.context || 'unknown'}"`
          )
        })
      }
    } else {
      console.error('❌ Unexpected error:', error)
    }
  }
}

generateDocs().catch(console.error)
