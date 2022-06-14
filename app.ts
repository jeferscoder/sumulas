import excel from 'excel4node';
import XLSX from "read-excel-file/node";


XLSX("sumulas.xlsx").then((rows) => {
    // `rows` is an array of rows
    // each row being an array of cells.
    
    console.log(rows.length)

    let di = []

    rows.map((r,i) => {
        di.push(rows[i][11])
    })
  
    let diarios = di.join('')
  

const protocolos = diarios.match(/\d{5}\/2020/g);


const diario = diarios.split(/\d{5}\/2020/g)



var workbook = new excel.Workbook()
var worksheet = workbook.addWorksheet('Sheet 1');

worksheet.cell(1,1).string('Protocolo')
worksheet.cell(1,6).string('SUMULAS')
worksheet.cell(1,2).string('CNPJ')
worksheet.cell(1,3).string('Nome')


protocolos?.map((data,index) => {
    worksheet.cell(2 + index,1).string(data)
})


diario.map((data,index) => {
    const self = {
        CNPJ:undefined,
        nome:undefined,
        sumulas:undefined,
    }
    const CNPJ = data.match(/[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}\/?[0-9]{4}\-?[0-9]{2}/g) as []
    const nome = data.match
    self.sumulas = data;
    if (CNPJ != null) {
        self.CNPJ = CNPJ[0];
    }

    if(data.includes('PRÉVIA')) {
        self.nome = data.slice(data.search('PRÉVIA') + 7,data.search('CNPJ'))
    }


    if(data.includes('SIMPLIFICADA') && data.includes('CNPJ')) {
        self.nome = data.slice(data.search('SIMPLIFICADA') + 13,data.search('CNPJ'))
    }

    if(data.includes('SIMPLIFICADA') && data.includes('torna público')) {
        let novo = data.slice(data.search('SIMPLIFICADA') + 13,data.search('torna público'))
        self.nome = novo.slice(0,novo.search('CNPJ'))
    }

    if(data.includes('OPERAÇÃO')) {
        self.nome = data.slice(data.search('OPERAÇÃO') + 2,data.search('CNPJ'))
    }

    if(data.includes('EMPRESARIAL')) {
        self.nome = data.slice(data.search('EMPRESARIAL') + 15,data.search('CNPJ'))
    }

    if(data.includes('OPERAÇÃO') && data.includes('torna público')) {
        const novo = data.slice(data.search('OPERAÇÃO') + 9,data.search('torna público'))
        self.nome = novo.slice(0,novo.search('CNPJ'))
    }

    if(data.includes('INSTALAÇÃO') && data.includes('torna público')) {
        const novo = data.slice(data.search('INSTALAÇÃO') + 9,data.search('torna público'))
        self.nome = novo.slice(0,novo.search('CNPJ'))
    }

    worksheet.cell(2 + index,6).string(self.sumulas)
    worksheet.cell(2 + index,3).string(self.nome)
    worksheet.cell(2 + index,2).string(self.CNPJ)

    /*
    if (data.includes('OPERAÇÃO A')) {
        self.nome = data.slice(data.search('OPERAÇÃO A') + 10,data.search('CNPJ'))
    } else

    if (data.includes('SIMPLIFICADA')) {
        self.nome = data.slice(data.search('SIMPLIFICADA') + 10,data.search('CNPJ'))
    } else {
        self.nome = data.slice(0,data.search('CNPJ'))
    }
    */
    //console.log(self)
})

workbook.write('Excel.xlsx')
console.log('processo finalizado')

});

/*
var workbook = new excel.Workbook()


var worksheet = workbook.addWorksheet('Sheet 1');

worksheet.cell(1,1).string('Protocolo')
worksheet.cell(1,6).string('SUMULAS')
worksheet.cell(1,2).string('CNPJ')
worksheet.cell(1,3).string('Seção')

protocolos?.map((data,index) => {
    worksheet.cell(2 + index,1).string(data)
})

const paginas = sumulas.split(/\d{5}\/2020/g)
const dir = diarios.split(/\d{5}\/2020/g) 


console.log(dir.length)


dir.map((data,index) => {
    worksheet.cell(2 + index,6).string(data)
})
*/
/*
paginas.map((data,index) => {
    worksheet.cell(2 + index,6).string(data)

    if (data.match('SÚMULA DE RECEBIMENTO DE RENOVAÇÃO')) {
        worksheet.cell(2+index,3).string('IAT- Súmulas de Renovação')
    }

    if (data.match('SÚMULA DE RECEBIMENTO DE LICENÇA')) {
        worksheet.cell(2+index,3).string('IAT - Súmulas de Recebimento')
    }

    if (data.match('SÚMULA DE REQUERIMENTO DE LICENÇA DE INSTALAÇÃO')) {
        workbook.cell(2+index,3).string('SÚMULA DE REQUERIMENTO DE LICENÇA DE INSTALAÇÃO')
    }

    const CNPJ = data.match(/[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}\/?[0-9]{4}\-?[0-9]{2}/g) as []
    if (CNPJ != null) {
        worksheet.cell(2+ index,2).string(CNPJ)
    }
})
*/

//workbook.write('Excel.xlsx');

