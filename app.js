const { createBot, createProvider, createFlow, addKeyword } = require('@bot-whatsapp/bot')
const QRPortalWeb = require('@bot-whatsapp/portal')
const BaileysProvider = require('@bot-whatsapp/provider/baileys')
const MockAdapter = require('@bot-whatsapp/database/mock')
const ExcelJS = require('exceljs')



// Flujos de chat
const flowSecundario = addKeyword(['siguiente']).addAnswer(['📄 Aquí tenemos el flujo secundario'])

const flow1 = addKeyword(['1']).addAnswer(
    [
        '📄 ',
    ],
    null,
    null,
    [flowSecundario]
)

const flow2 = addKeyword(['2']).addAnswer(
    [
        '🙌 ',
    ],
    null,
    null,
    [flowSecundario]
)

const flow3 = addKeyword(['3']).addAnswer(
    [
        '🚀 ',
    ],
    null,
    null,
    [flowSecundario]
)

const flow4 = addKeyword(['4']).addAnswer(
    ['🤪 '],
    null,
    null,
    [flowSecundario]
)

const flowPrincipal = addKeyword(['Mantenimiento'])
    .addAnswer('🙌 Hola bienvenido a este *Chatbot*')
    .addAnswer(
        [
            'te comparto los siguientes links',
            '👉 *1',
            '👉 *2',
            '👉 *3',
            '👉 *4'            

        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)
            
        },
        [flow1, flow2, flow3, flow4]
        
    )



    const saveExcel = async (data) => {

        const workbook = new ExcelJS.Workbook();
  
        // Intentar leer el archivo existente
        const fileName = 'Registros2.xlsx';
        let worksheetName = 'Registros2';
        try {
            await workbook.xlsx.readFile(fileName);
        } catch (error) {
            console.log('El archivo no existe, se creará uno nuevo.');
        }

        // Obtener la hoja de cálculo o crear una nueva si no existe
        let sheet = workbook.getWorksheet(worksheetName);
        if (!sheet) {
            sheet = workbook.addWorksheet(worksheetName);
            sheet.columns = [
                { header: 'Area', key: 'area', width: 25 },
            ];
        }
      
        const lastRowNumber = sheet.lastRow ? sheet.lastRow.number : 0;
        const newRowNumber = lastRowNumber + 1;
        const newRow = sheet.getRow(newRowNumber);
        newRow.values = [data];
        newRow.commit();

        // Añadir nuevas filas
        //sheet.addRows([{area:data}]);
    
        // Guardar el archivo
        workbook.xlsx.writeFile(fileName)
        .then(() => {
            console.log('Guardado o actualizado satisfactoriamente');
        })
        .catch((error) => {
            console.error('Error al guardar o actualizar el archivo:', error);
        });
    }
    

const main = async () => {
    let data = []
    const adapterDB = new MockAdapter()
    const adapterFlow = createFlow([flowPrincipal])
    const adapterProvider = createProvider(BaileysProvider)

    createBot({
        flow: adapterFlow,
        provider: adapterProvider,
        database: adapterDB,
    })
    data.push(
        {
            area: flowPrincipal
        }
    )

    QRPortalWeb()
    
}






main()
