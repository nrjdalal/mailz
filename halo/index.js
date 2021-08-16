/*

# transport information w/o defaults
USERNAME=
USERPASS=

# sender information w/o defaults
FROM_NAME=
FROM_MAIL=

# transport information with defaults (smtp.gmail.com, 465, true)
HOST=
PORT=
SECURE=

# delay between emails in ms (default = 5000)
DELAY=

# receivers (single address, comma seperated values, list file or xlsx file)
TO_MAIL=

# subject as a string (has server test default)
SUBJECT=

# content as a filename or string (has server test default)
HTML=

*/

require('dotenv').config()
const fs = require('fs')
const nodemailer = require('nodemailer')
const validate = require('validate.js')
const xlsx = require('xlsx')

// 1. helper utilities
// 1.1 email validator
const isEmail = (email) => {
  const res = validate({ from: email }, { from: { email: true } })
  if (res) {
    return false
  } else {
    return true
  }
}
// 1.2 delay timer
const timer = (ms) => new Promise((res) => setTimeout(res, ms))

// 2. nodemailer configuration
const sender = async (content, to) => {
  let transporter = nodemailer.createTransport({
    host: process.env.HOST || 'smtp.gmail.com',
    port: process.env.PORT || 465,
    secure: process.env.SECURE || true,
    auth: {
      user: process.env.USERNAME,
      pass: process.env.USERPASS,
    },
  })

  await transporter.sendMail({
    from: `${process.env.FROM_NAME} <${process.env.FROM_MAIL}>`,
    to: to,
    subject: process.env.SUBJECT || 'This is a Test Mail',
    html: content || 'The given credentials are working correctly!',
  })

  console.log('Email sent to ' + to)
}

// 3. mailer (one-by-one)
const mailer = async (content, to, delay) => {
  // email sender
  fs.readFile(process.cwd() + `/${to}`, 'utf8', (err, res) => {
    let index = 1
    if (err) {
      to = to.replace(/ /g, '')
      // if empty
      if (to.length === 0) {
        sender(content, process.env.FROM_MAIL)
        return
      }
      // to single email address
      if (isEmail(to)) {
        sender(content, to)
      }
      // to comma separated emails
      if (to.includes(',')) {
        async function sendmail() {
          for (x of to.split(',')) {
            console.log(index++)
            if (isEmail(x)) {
              sender(content, x)
              await timer(delay)
            } else {
              console.log('Invalid email address encountered - ' + x)
            }
          }
        }
        sendmail()
      }
    } else {
      if (to.includes('xlsx')) {
        let workbook = xlsx.readFile(to)
        for (sheet of workbook.SheetNames) {
          let worksheet = workbook.Sheets[sheet]
          async function sendmail() {
            for (x of Object.keys(worksheet)) {
              console.log(index++)
              if (typeof worksheet[x].v !== 'undefined') {
                res = worksheet[x].v.replace(/ /g, '')
                if (isEmail(res)) {
                  sender(content, res)
                  await timer(delay)
                } else {
                  console.log('Invalid email address encountered - ' + x)
                }
              }
            }
          }
          sendmail()
        }
      } else {
        res = res.replace(/ /g, '')
        // list file
        async function sendmail() {
          for (x of res.split('\n')) {
            console.log(index++)
            if (isEmail(x)) {
              sender(content, x)
              await timer(delay)
            } else {
              console.log('Invalid email address encountered - ')
            }
          }
        }
        sendmail()
      }
    }
  })
}

// operating function
fs.readFile(process.cwd() + `/${process.env.HTML}`, 'utf8', (err, res) => {
  if (err) {
    // email as string
    mailer(process.env.HTML, process.env.TO_MAIL, process.env.DELAY || 5000)
  } else {
    // email as html file
    mailer(res, process.env.TO_MAIL, process.env.DELAY || 5000)
  }
})
