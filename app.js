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
            
        },
        [flow1, flow2, flow3, flow4]
        
    )



    const saveExcel = (data) => {

        const workbook = new ExcelJS.Workbook()
        const fileName = 'Registros.xlsx'
        const sheet = workbook.addWorksheet('Registros')
        const reColumns = [
            { Headers: 'Area', key: 'area'}           
        ]    
        sheet.columns = reColumns
        sheet.addRows(data)
        workbook.xlsx.writeFile(fileName).then((e) => {
            console.log('Creado satisfatoriamente');
        })
        .catch(() => {
            console.error("Error al crear el archivo");
        })    
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
    saveExcel(data)
}






main()

