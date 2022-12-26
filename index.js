const reader = require('xlsx')
const fs = require('fs/promises')
const fssync = require('fs')
const prompt = require('prompt-sync')({sigint: true});



const path = prompt("Enter the path of the file: ")
async function readXlsxFile(path) {
    let data = []
    let result = {}
    const workbook = reader.readFile(path)
    const sheets = workbook.SheetNames
  
    for (let i = 0; i < sheets.length; i++) {
        const temp = reader.utils.sheet_to_json(
            workbook.Sheets[workbook.SheetNames[i]])
        temp.forEach((res) => {
            data.push(res)
        })
    }
  
    // This is for csv file
    if (path.endsWith('.csv')) {
        const data = await fs.readFile(path, 'utf8')
    }
    
    for (let i = 0; i < data.length; i++) {
        if (data[i].hasOwnProperty('Amount')) {
            if (result[data[i]['Customer Id']]) {
                result[data[i]['Customer Id']].EndingBalance += data[i]['Amount']

                if (result[data[i]['Customer Id']].MinBalance > data[i]['Amount']) {
                    result[data[i]['Customer Id']].MinBalance = data[i]['Amount']
                }
                if (result[data[i]['Customer Id']].MaxBalance < data[i]['Amount']) {
                    result[data[i]['Customer Id']].MaxBalance = data[i]['Amount']
                }
            }
            else {
                result[data[i]['Customer Id']] = {
                    CustomerID: data[i]['Customer Id'], 'MM/YYYY':new Date().getMonth(11) + '/' + new Date().getFullYear(2022),
                    MinBalance: data[i]['Amount'], MaxBalance: data[i]['Amount'], EndingBalance: data[i]['Amount']
                }
            }

        }

    }

    let resultArray = Object.values(result).map((item, )=> Object.values(item))

    const rows = ['CustomerID', 'MM/YYYY', 'MinBalance', 'MaxBalance', 'EndingBalance\r\n'].join(",") + resultArray.join("\r\n")
    const outputPath = `${path.slice(0, path.lastIndexOf("."))}result.csv`

    if ( fssync.existsSync(outputPath)) {
        const overwrite = prompt(`File already exists at ${outputPath} Press yes to overwrite the file`)
        if (overwrite.toLowerCase() !== 'yes') {
          return console.log('File not created')
        }
    }
        fs.writeFile(outputPath, rows, (err) => {
            if (err)
                console.error(err)
            else
                console.log(`File has been created at ${outputPath}`)
        })

    


 }
readXlsxFile(path)
    

