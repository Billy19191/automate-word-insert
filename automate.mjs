import fs from 'fs'
import path from 'path'
import ExcelJS from 'exceljs'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import libre from 'libreoffice-convert'

const templatePath = 'input/template.docx'
const outputDir = 'output'
const excelPath = 'input/companyList.xlsx'

if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir)

async function generateDocs() {
  try {
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.readFile(excelPath)
    const worksheet = workbook.getWorksheet(1)

    const templateContent = fs.readFileSync(templatePath, 'binary')
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

        const docBuffer = doc.getZip().generate({ type: 'nodebuffer' })
        const docxFilename = `${companyNumber}_${companyInitial} Schedule 9_CC.docx`
        const docxPath = path.join(outputDir, docxFilename)
        fs.writeFileSync(docxPath, docBuffer)

        try {
          // Convert to promise-based usage
          const pdfBuffer = await new Promise((resolve, reject) => {
            libre.convert(docBuffer, '.pdf', undefined, (err, data) => {
              if (err) {
                reject(err)
              } else {
                resolve(data)
              }
            })
          })

          const pdfFilename = `${companyNumber}_${companyInitial} Schedule 9_CC.pdf`
          const pdfPath = path.join(outputDir, pdfFilename)
          fs.writeFileSync(pdfPath, pdfBuffer)
          console.log(`📄 Generated PDF: ${pdfFilename}`)
          successCount++
        } catch (pdfError) {
          console.error(
            `❌ PDF conversion failed for ${companyNumber}:`,
            pdfError.message
          )
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
