const { createBot, createProvider, createFlow, addKeyword } = require('@bot-whatsapp/bot')
const QRPortalWeb = require('@bot-whatsapp/portal')
const BaileysProvider = require('@bot-whatsapp/provider/baileys')
const MockAdapter = require('@bot-whatsapp/database/mock')
const ExcelJS = require('exceljs')
const nodemailer = require('nodemailer')


let area;
let nombreSolicitante;
let nombreArea;
let descripcionMotivo;
let prioridad;


// Flujos de chat
const flow1 = addKeyword(['1'])
  .addAnswer('Nombre del solicitante',
    { capture: true }, (ctx) => {
      console.log('Mensaje entrante:', ctx.body);      
      nombreSolicitante = ctx.body;
    })
  .addAnswer(
    [
      '📄 Infraestructura ',
      'Escribe el nombre del área'
    ],
    { capture: true }, (ctx) => {
      console.log('Mensaje entrante:', ctx.body);      
      nombreArea = ctx.body;
    })
  .addAnswer('Escribe un breve descripción del motivo',
    { capture: true }, (ctx) => {
      console.log('Mensaje entrante:', ctx.body);     
      descripcionMotivo = ctx.body;
    })
  .addAnswer(['📄 Dale un valor de prioridad',
    '👉 *1 Alta   Equipo o área sin funcionamiento',
    '👉 *2 Media  Equipo o área funcional pero con restricciones',
    '👉 *3 Baja   El equipo o área necesitan una inspección'
  ],
    { capture: true }, (ctx) => {
      console.log('Mensaje entrante:', ctx.body);
      saveExcel(ctx.body);
      correoEnviado(ctx.body);
      prioridad = ctx.body;
    })
  .addAnswer("Tu solicitud ha sido recibida, ¡gracias!");


const flow2 = addKeyword(['2'])
    .addAnswer('Nombre del solicitante',
            { capture: true }, (ctx) => {  
                console.log('Mensaje entrante:', ctx.body) 
                nombreSolicitante =ctx.body;    
        })
    .addAnswer(
        [
            '🙌 Maquinas y equipos',
            'Escribe el nombre del equipos'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            nombreArea =ctx.body;
            
                   
    })
    .addAnswer('escribe un breve descripcion del motivo',
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
           descripcionMotivo = ctx.body; 
                       
    }) 
    .addAnswer(['📄 Dale un valor de prioridad',    
            '👉 *1 Alta   Equipo o area sin funcionamiento',
            '👉 *2 Media  Equipo o area funcional pero con restrinciones',
            '👉 *3 Baja   El quipo o area necesecitan un inspeccion'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            prioridad =ctx.body
            saveExcel(ctx.body)
            correoEnviado(ctx.body)                    
    })
    .addAnswer("Tu solicitud ha sido recibida, ¡gracias!")           

const flow3 = addKeyword(['3'])
.addAnswer('Nombre del solicitante',
    { capture: true }, (ctx) => {  
        console.log('Mensaje entrante:', ctx.body) 
        nombreSolicitante = ctx.body;   

    })
    .addAnswer(
        [
            '🚀 Sistemas ',
            'Escribe el nombre del equipos'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            nombreArea = ctx.body;
                  
    })
    .addAnswer('escribe un breve descripcion del motivo',
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            descripcionMotivo = ctx.body;
                        
    }) 
    .addAnswer(['📄 Dale un valor de prioridad',    
            '👉 *1 Alta   Equipo o area sin funcionamiento',
            '👉 *2 Media  Equipo o area funcional pero con restrinciones',
            '👉 *3 Baja   El quipo o area necesecitan un inspeccion'
        ],
        { capture: true }, (ctx) => {            
            console.log('Mensaje entrante:', ctx.body)
            saveExcel(ctx.body) 
            correoEnviado(ctx.body);
            prioridad = ctx.body;                   
    })
    .addAnswer("Tu solicitud ha sido recibida, ¡gracias!")       
    

    const flowPrincipal = addKeyword(['Mantenimiento'])
    .addAnswer('🙌 Hola bienvenido a este Chatbot de matenimiento')
    .addAnswer(
      [
        'selecciona el area de necesidad ',
        '👉 *1 Infraestructura',
        '👉 *2 Maquninas y equipos',
        '👉 *3 Sistemas',
      ],
      {capture: true},
      (ctx1) => {
        const respuestas = ctx1.body.split(' ');
        const respuestasValidas = respuestas.filter(element => element.match(/[123]/));
        if (respuestasValidas.length > 0) {
          console.log('Mensaje entrante:', respuestasValidas);  
          area =ctx1.body        
        } else {
          console.log('Respuesta no válida');
        }
      },
      [flow1, flow2, flow3]
    );


 
        
// configuracion de excel.
const saveExcel = async (data) => {
    const workbook = new ExcelJS.Workbook();  
    // Intentar leer el archivo existente
    const fileName = 'Registros.xlsx';
    let worksheetName = 'Registros';      
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
            { header: 'Fecha', key: 'fecha', width: 25 },
            { header: 'Area', key: 'area', width: 25 },
            { header: 'Nombre', key: 'nombre', width: 25 },           
            { header: 'Equipo', key: 'equipo', width: 25 },
            { header: 'Motivo', key: 'motivo', width: 25 },
            { header: 'Prioridad', key: 'prioridad', width: 25 }                                              
    ];                        
    }     
    const lastRowNumber = sheet.lastRow ? sheet.lastRow.number : 0;
    const newRowNumber = lastRowNumber + 0;
    const newRow = sheet.getRow(newRowNumber);    
    var dat= new Date();
    sheet.addRow([dat,area,nombreSolicitante,nombreArea,descripcionMotivo,prioridad]);  
    newRow.commit();     
    workbook.xlsx.writeFile(fileName)           
    .then(() => {
        console.log('Guardado o actualizado satisfactoriamente');
    })
    .catch((error) => {
        console.error('Error al guardar o actualizar el archivo:', error);
    });
}

// correo electrónico
const correoEnviado = async (data) => {
    
let transporter = nodemailer.createTransport({
    service: 'zoho',
    auth: {
        user: 'mantenimiento@intercalco.com',
        pass: 'Intercalco*'
    }
});

// Configurar el correo electrónico a enviar
let mailOptions = {
    from: 'mantenimiento@intercalco.com',
    to: 'mantenimiento@intercalco.com',
    subject: nombreSolicitante,
    text: 
    `${nombreArea}\n${descripcionMotivo}`
        
};

// Enviar el correo electrónico
transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
        console.log('Error:', error);
    } else {
        console.log('Email sent:', info.messageId);
    }
});
}

  

// constante 
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