import xlsx, { Row } from "read-excel-file/node";
import excel from "excel4node";
import cluster from 'cluster';
import os from 'os';
import fs from 'fs';
import { match } from "assert";
// /SÚMULA(.*?)\d{5}\/2020/g

const numOfCpuCores = os.cpus().length

const app = async () => {
  

  const planilhas = await xlsx(fs.createReadStream("./input/sumulas.xlsx"));
 const rows = []

 console.log("selecionando linhas da coluna PAGINA DIARIO...")
 planilhas.map((row,index) => {
  rows.push(planilhas[index][11])
 })

  /*
 planilhas.map(row => {
  console.log(row)
  rows.push(row[11])
 })*/

 console.log('juntando todas as linhas...')

 const uniques = [...new Set(rows)].join('')


 console.log('separando as sumulas')
 const sumulas = uniques.split(/\d{5}\/2020/g)
 const protocolos = uniques.match(/\d{5}\/2020/g)

 //const sumulas = uniques.match()
 
 var workbook = new excel.Workbook()
 var worksheet = workbook.addWorksheet('Sheet 1');
 worksheet.cell(1,1).string('Protocolos')
 worksheet.cell(1,2).string('CNPJ')
 worksheet.cell(1,3).string('Local')
 worksheet.cell(1,4).string('Nomes')
 worksheet.cell(1,5).string('Atividade')
 worksheet.cell(1,6).string('Edição')
 worksheet.cell(1,7).string('Súmulas')


 console.log('preparando planilha')

 console.log('alocando protocolos')
 protocolos?.map((pro,index) => {
  worksheet.cell(2 + index,1).string(pro)
 })

 console.log('alocando cnpj, edicao, nome, local ,sumulas')

 let edit = ''
 sumulas.map((sum,index) => {
  const cell = {
    cnpj:'',
    edicao:'',
    local:'',
    sumulas:'',
    nomes:'',
    atividade:''
  }

  

  if (sum.match(/Edição\snº\s\d{5}/s)) {
    edit = sum.match(/Edição\snº\s\d{5}/s)[0]
  }


  cell.sumulas = sum
  cell.cnpj =  sum.match(/[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}\/?[0-9]{4}\-?[0-9]{2}/g)
  cell.local = sum.match(/(implantada | instalada).*/g)


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


  if (sum.match('Prévia')) {
    if (sum.match('a ser implantada')) {
      cell.atividade = sum.match(/(?<=Prévia para).*?(?=a ser implantada)/g)
    } else {
      cell.atividade = sum.match(/(?<=Prévia).*?(?=situada)/g)
    }

    
  }

  if (sum.match('Instalação')) {
    cell.atividade = sum.match(/(?<=Instalação para).*?(?=a ser implantada)/g)
  }

  if (sum.match('Operação')) {
    cell.atividade = sum.match(/(?<=Operação para).*?(?=Licença)/g) || sum.match(/(?<=Operação para).*?(?=instalada)/g) || 
    sum.match(/(?<=Operação).*?(?=situada)/g)
  }

  if (sum.match('Simplificada')) {

    if (sum.match('a ser implantada')) {
    cell.atividade = sum.match(/(?<=Simplificada para).*?(?=a ser implantada)/g)
    } else {
      cell.atividade = sum.match(/(?<=Simplificada para).*?(?=implantada)/g) 
    }

    
  }

  if (sum.match('Regularização')) {
    cell.atividade = sum.match(/(?<=Regularização para).*?(?=instalada)/g)
  }

  if (sum.match('Ambiental')) {
    cell.atividade = sum.match(/(?<=Regularização para).*?(?=instalada)/g)
  }

  /*
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

  if (sum.match('Licença Simplificada')) {
    cell.atividade = 'Recebimento de Licença Simplificada - CIS'
  }



  if (sum.match('Licença Simplificada')) {
    cell.atividade = 'Recebimento de Licença Simplificada - CIS'
  }
  */

  worksheet.cell(2 + index,2).string(cell.cnpj)
  worksheet.cell(2 + index,3).string(cell.local)
  worksheet.cell(2 + index,4).string(cell.nomes)
  worksheet.cell(2 + index,5).string(cell.atividade)
  worksheet.cell(2 + index,6).string(edit)
  worksheet.cell(2 + index,7).string(cell.sumulas)

  console.log(cell)

 })

  workbook.write('./output/Excel.xlsx');
  console.log('finalizado...')
};

app();
