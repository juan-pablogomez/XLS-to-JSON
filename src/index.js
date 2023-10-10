const XLSX = require("xlsx")
const FS = require("fs")

const ExcelAJson = () => {
    const excel = XLSX.readFile(
        "C:\\Users\\ASUS\\OneDrive - culturantioquia.gov.co\\Caracterizacion\\XLS a JSON\\organizaciones_01.xlsx"
    )
    var hoja = excel.SheetNames;
    let datos = XLSX.utils.sheet_to_json(excel.Sheets[hoja[0]])

    const jsonDatos = []

    for( let i=0; i < datos.length; i++) {
        const dato = datos[i]
        jsonDatos.push({
            ...dato, 
            register_date: new Date(((dato.register_date - (25567 + 2)) * 86400 * 1000)) 
        })
    }


    // console.log(Buffer.from(JSON.stringify(jsonDatos)).toString())
    
    // Funcion writeFile con nombre de Archivo, contenido y callback function en argumentos
    FS.writeFile('organizaciones_01.json', Buffer.from(JSON.stringify(jsonDatos)).toString(), function (err) { 
        if (err) throw err
        console.log('El Archivo quedo Melo')
    })


    // Funcion writeFile con nombre de Archivo, contenido y callback function en argumentos
    // FS.writeFileSync('organizaciones_01.json', Buffer.from(JSON.stringify(jsonDatos)).toString())

}

ExcelAJson()