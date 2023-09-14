const { createBot, createProvider, createFlow, addKeyword } = require('@bot-whatsapp/bot')
const QRPortalWeb = require('@bot-whatsapp/portal')
const BaileysProvider = require('@bot-whatsapp/provider/baileys')
const MockAdapter = require('@bot-whatsapp/database/mock')
const ExcelJS = require('exceljs')

// Flujos de chat
const flow1 = addKeyword(['1'])
    .addAnswer(
        [
            '📄 Infraestructura ',
            'Escribe el nombre del area'
        ],    
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)
    })
    .addAnswer('escribe un breve descripcion del motivo',
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)              
    }) 
    .addAnswer(['📄 Dale un valor de prioridad',    
            '👉 *1 Alta   Equipo o area sin funcionamiento',
            '👉 *2 Media  Equipo o area funcional pero con restrinciones',
            '👉 *3 Baja   El quipo o area necesecitan un inspeccion'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)                   
    })    

const flow2 = addKeyword(['2'])
    .addAnswer(
        [
            '🙌 Maquinas y equipos',
            'Escribe el nombre del equipos'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)        
    })
    .addAnswer('escribe un breve descripcion del motivo',
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)              
    }) 
    .addAnswer(['📄 Dale un valor de prioridad',    
            '👉 *1 Alta   Equipo o area sin funcionamiento',
            '👉 *2 Media  Equipo o area funcional pero con restrinciones',
            '👉 *3 Baja   El quipo o area necesecitan un inspeccion'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)                    
    })        

const flow3 = addKeyword(['3'])
    .addAnswer(
        [
            '🚀 Sistemas ',
            'Escribe el nombre del equipos'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)        
    })
    .addAnswer('escribe un breve descripcion del motivo',
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)              
    }) 
    .addAnswer(['📄 Dale un valor de prioridad',    
            '👉 *1 Alta   Equipo o area sin funcionamiento',
            '👉 *2 Media  Equipo o area funcional pero con restrinciones',
            '👉 *3 Baja   El quipo o area necesecitan un inspeccion'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)                    
    })    
   
const flowPrincipal = addKeyword(['Mantenimiento'])
    .addAnswer('🙌 Hola bienvenido a este *Chatbot de matenimiento*')
    .addAnswer(
        [ 'selecciona el area de necesidad ',
            '👉 *1 Infraestructura',
            '👉 *2 Maquninas y equipos',
            '👉 *3 Sistemas'
        ],
        {capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body)                    
        },
        [flow1, flow2, flow3] 
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
                { header: 'Fecha.', key: 'fecha', width: 10, style: { numFmt: 'dd/mm/yyyy' } },
                { header: 'Area', key: 'area', width: 25 },
                { header: 'Equipo', key: 'equipo', width: 25 }
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
    
    QRPortalWeb()    
}
main()