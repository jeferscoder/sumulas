import xlsx, { Row } from "read-excel-file/node";
import excel from "excel4node";
import fs from 'fs';
// /SÚMULA(.*?)\d{5}\/2020/g
const app = async () => {
  

  const planilhas = await xlsx(fs.createReadStream("./input/sumulas.xlsx"));
 const rows = []

 for (let i = 0;i < planilhas.length;i++) {
  rows[i] = planilhas[i][11]
 }

  /*
 planilhas.map(row => {
  console.log(row)
  rows.push(row[11])
 })*/

 console.log('filtrando sumulas repetidas')
 const uniques = [...new Set(rows)].join('')

///(((Edição\snº\s\d{5})|| SÚMULA)(.*?)\d{5}\/2020)/g
 const sumulas = uniques.match(/(.*?)\d{5}\/2020/g)
 
 var workbook = new excel.Workbook()
 var worksheet = workbook.addWorksheet('Sheet 1');
 worksheet.cell(1,1).string('Protocolos')
 worksheet.cell(1,2).string('CNPJ')

 worksheet.cell(1,4).string('Nomes')
 worksheet.cell(1,5).string('Local')
 worksheet.cell(1,6).string('Edição')
 worksheet.cell(1,7).string('Súmulas')


 sumulas.map((sum,index) => {
  const cell = {
    protocolos:'',
    cnpj:'',
    edicao:'',
    sumulas:'',
    nomes:'',
    atividade:'',
    local:''
  }

  if (sum.match(/Edição\snº\s\d{5}/s) != null) {
    cell.edicao = sum.match(/Edição\snº\s\d{5}/s)[0]
  }

  cell.protocolos = sum.match(/\d{5}\/2020/g)
  cell.sumulas = sum
  cell.cnpj =  sum.match(/[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}\/?[0-9]{4}\-?[0-9]{2}/g)


  if (sum.match('INSTALAÇÃO')) {
    const current = sum.slice(sum.search('INSTALAÇÃO') + 11,sum.search('torna público')).slice(0)
   cell.nomes = current.slice(0,current.search('CNPJ'))
  } else 

  if (sum.match('OPERAÇÃO')) {
    const current =  sum.slice(sum.search('OPERAÇÃO') + 9,sum.search('torna público'))
   cell.nomes = current.slice(0,current.search('CNPJ'))
  }


  if (sum.match('PRÉVIA')) {
    const current = sum.slice(sum.search('PRÉVIA') + 7,sum.search('torna público'))
   cell.nomes = current.slice(0,current.search('CNPJ'))
  }

  if (sum.match('SIMPLIFICADA')) {
    const current = sum.slice(sum.search('SIMPLIFICADA') + 13,sum.search('torna público'))
    cell.nomes = current.slice(0,current.search('CNPJ'))
  }


  if (sum.match('Licença Prévia')) {
    cell.atividade = 'Recebimento de Licença Prévia - CIS'
  }

  if (sum.match('Renovação de Licença de Operação')) {
    cell.atividade = 'Recebimento de Renovação de Licença de Operação - CIS'
  }

  if (sum.match('Licença de Operação')) {
    cell.atividade = 'Recebimento de Licença de Operação - CIS'
  }

  if (sum.match('Licença Prévia')) {
    cell.atividade = 'Recebimento de Licença Prévia - CIS'
  }

  if (sum.match('Renovação de Licença Simplificada')) {
    cell.atividade = 'Requerimento de Renovação de Licença Simplificada - CIS'
  }

  if (sum.match('Autorização Ambiental')) {
    cell.atividade = 'Recebimento de Autorização Ambiental - CIS'
  }

  if (sum.match('Renovação de Licença Simplificada')) {
    cell.atividade = 'Requerimento de Renovação de Licença Simplificada - CIS'
  }


 
  worksheet.cell(2 + index,1,2).string(cell.cnpj)
  worksheet.cell(2 + index,1,3).string(cell.atividade)
  worksheet.cell(2 +  index,1,4).string(cell.nomes)
  worksheet.cell(2 + index,1,5).string(cell.local)
  worksheet.cell(2 + index,1,6).string(cell.edicao)
  worksheet.cell(2 + index,1,7).string(cell.sumulas)


 })

  workbook.write('./output/Excel.xlsx');
  console.log('finalizado...')
};

app();
