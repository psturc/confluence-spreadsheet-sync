const async = require('async')
const cheerio = require('cheerio')
const google = require('googleapis')
const request = require('request')
const winston = require('winston')
const YAML = require('yamljs')

const auth = require('./auth')
// Disable verification of self-signed certificate
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0'

var requirements = {}
var sheetArray = []

const config = YAML.load('config.yaml')

async.parallel({
  oauth2Client: (parallelCb) => {
    auth((err, oauth2Client) => {
      if (err) {
        return parallelCb(`Error when requesting oauth2Client. ${err}`)
      }
      parallelCb(null, oauth2Client)
    })
  },
  htmlContent: (parallelCb) => {
    request.get(config.confluence_page_url, (err, res, html) => {
      if (err) {
        return parallelCb(`Error when requesting the content from page ${config.confluence_page_url}. ${err}`)
      }
      parallelCb(null, html)
    })
  }

}, async (err, res) => {
  if (err) {
    return winston.error(err)
  }
  const oldSheetData = await getGoogleSheetData(res.oauth2Client)
  filterHtmlContent(res.htmlContent)
  
  generateUpdatedSheetContent(oldSheetData)
  updateGoogleSheet(res.oauth2Client)
})

function filterHtmlContent(html) {
  let $ = cheerio.load(html)
  $('table tbody tr').each(function(){
    let title = $(this).find('a').text() 
    let status = $(this).find('.status-macro').text()
    requirements[title] = status
  })
}

function generateUpdatedSheetContent(rows) {
  for (let i = 0; i < rows.length; i++) {
    let row = rows[i]
    sheetArray.push([ requirements[row[0]] ])
  }
}

async function getGoogleSheetData(auth) {
  return new Promise((resolve, reject) => {
    let sheets = google.sheets('v4')
    sheets.spreadsheets.values.get({
      auth: auth,
      spreadsheetId: config.spreadsheet.id,
      range: `${config.spreadsheet.sheet_name}!${config.spreadsheet.column_keys}`,
    }, function(err, response) {
      if (err) {
        reject('The API returned an error: ' + err)
        return
      }
      let rows = response.values
      if (rows.length == 0) {
        reject('No data found.')
      } else {
        resolve(rows)
      }
    })
  })
}

function updateGoogleSheet(auth) {
  let body = {
    values: sheetArray
  }
  let sheets = google.sheets('v4')
  sheets.spreadsheets.values.update({
    auth: auth,
    valueInputOption: 'USER_ENTERED',
    spreadsheetId: config.spreadsheet.id,
    range: `${config.spreadsheet.sheet_name}!${config.spreadsheet.column_values}`,
    resource: body
  }, function(err, result) {
    if (err) {
    // Handle error
      winston.error(err)
    } else {
      winston.info('%d cells updated.', result.updatedCells)
    }
  })
}