const nodemailer = require('nodemailer')
const ExcelJS = require('exceljs')
const { join } = require('path')

module.exports = async (req, res) => {
  const { coaching_enrollee_xlsx } = JSON.parse(req.body.data)
  console.log('1) Zoho object parsed into JS object');

  const coachName = coaching_enrollee_xlsx.shift()

  const coachingWb = new ExcelJS.Workbook()
  const coachingWs = await coachingWb.xlsx.readFile(join(__dirname, '_files', 'CoachingSpreadsheet.xlsx'))
  const enrolleesSheet = coachingWs.getWorksheet('Enrollees')
  const START_OF_ENROLLEE_LIST = 3
  coaching_enrollee_xlsx.forEach((enrolleeDetails, index) => {
    enrolleesSheet.getCell(`A${index + START_OF_ENROLLEE_LIST}`).value = index + 1
    enrolleesSheet.getCell(`B${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.workshopDate
    enrolleesSheet.getCell(`C${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.enrolleeName
    enrolleesSheet.getCell(`D${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.company
    enrolleesSheet.getCell(`E${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.address
    enrolleesSheet.getCell(`F${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.phone
    enrolleesSheet.getCell(`G${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.email
    enrolleesSheet.getCell(`H${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.course
    enrolleesSheet.getCell(`I${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.version
    enrolleesSheet.getCell(`J${index + START_OF_ENROLLEE_LIST}`).value = enrolleeDetails.type
  });
  console.log('2) Created coaching spreadsheet');

  try {
    //* 3) Create buffer
    const buffer = await facilitatorWb.xlsx.writeBuffer()
    //* 4) Create reusable transporter object using the default SMTP transport
    const transporter = nodemailer.createTransport({
      host: 'smtp.zoho.com',
      port: 465,
      secure: true, // true for 465, false for other ports
      auth: {
        user: 'brett.handley@prioritymanagement.com.au',
        pass: process.env.EMAIL_PW,
      },
    })
    const today = new Date()
    const sevenDaysAgo = new Date(today)
    sevenDaysAgo.setDate(today.getDate() - 7)
    //* 5) Send mail with defined transport object
    const emailRes = await transporter.sendMail({
      from: `'Priority Management Sydney' <brett.handley@prioritymanagement.com.au>`,
      to: 'johncodeinaire@gmail.com',
      subject: `Spreadsheet for coach - ${coachName}`,
      text: `Dear PM Admin,/r This is the spreadsheet for coach ${coachName}. It lists all the enrollees they've been assigned to coach from ${sevenDaysAgo.toDateString()} to ${today.toDateString()}/r Regards, PM Automation`,
      html: `<p>Dear PM Admin,</p><p>This is the spreadsheet for coach ${coachName}. It lists all the enrollees they've been assigned to coach from ${sevenDaysAgo.toDateString()} to ${today.toDateString()}</p><p>Regards, PM Automation</p>`,
      attachments: [
        {
          filename: `Coaching Spreadsheet - ${coachName}.xlsx`,
          content: buffer
        }
      ]
    })
    console.log('Message sent:', emailRes.messageId)
    res.json({body: `Message sent: ${emailRes.messageId}`})
  } catch (error) {
    console.error('Error:', error)
    res.json({body: `Error: ${error}`})
  }

}