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


const diarios = diarios.split(/\d{5}\/2020/g)



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




const diarios = rows[50][11];

const datas = [];

rows.map((data, index) => {
  datas.push(data[11]);
});

const sumulas =  [...new Set(datas)]

const protocolos = sumulas.join('').match(/\d{5}\/2020/g)


/*
const sumulas = diarios.toString().split(/\d{5}\/2020/g)

const protocolos = diarios.toString().match(/\d{5}\/2020/g)
  */

//console.log(protocolos)

/*
const datas = ["sc"];

rows.map((data, index) => {
  datas.push(data[9].toString());
});


const protocolos = [...new Set(datas.join("").match(/\d{5}\/2020/g))]

const sumulas = [...new Set(datas.join("").split(/\d{5}\/2020/g))]

  console.log(sumulas)

//console.log(sumulas)
/*


worksheet.cell(1, 2).string("Sumulas");
sumulas?.map((data, index) => {
  worksheet.cell(2 + index, 2).string(data);

});
*/
/*
const protocolos = sumulas.join("").match(/\d{5}\/2020/g);

const uniques = [...new Set(protocolos)]

const datas = []



worksheet.cell(1, 1).string("Protocolos");
uniques.map((data, index) => {
  //worksheet.cell(2 + index, 1).string(data);
  //console.log(data)
});
*/

workbook.write('./output/Excel.xlsx')
console.log("xlsx");


import xlsx, { Row } from "read-excel-file/node";
import excel from "excel4node";
// /SÚMULA(.*?)\d{5}\/2020/g
const app = async () => {
  const rows = await xlsx("./input/sumulas.xlsx");


  var workbook = new excel.Workbook()
  var worksheet = workbook.addWorksheet('Sheet 1');

  const cell = {
    protocolos:'',
    cnpj:'',
    edicao:'',
    local:'',
    sumulas:''
  }

  let counter = 0;

  worksheet.cell(1,1).string('Protocolos')
  worksheet.cell(1,2).string('CNPJ')
  worksheet.cell(1,3).string('Edição')
  worksheet.cell(1,4).string('Nomes')
  worksheet.cell(1,4).string('Local')
  worksheet.cell(1,5).string('Súmulas')


  const datas = []

  rows.map(row => datas.push(row[11]))

  const diarios = [...new Set(datas.join('').match(/(Edição\snº\s\d{5} | SÚMULA)(.*?)\d{5}\/2020/g))]
  
  diarios.map((sumulas,index) => {
    
    let edit = sumulas.match(/Edição\snº\s\d{5}/g)
    let protocolos = sumulas.match(/\d{5}\/2020/g)[0]
    let cnpj = sumulas.match(/[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}\/?[0-9]{4}\-?[0-9]{2}/g) as []
    let nomes = ''

    if (sumulas.match('INSTALAÇÃO')) {
      const current = sumulas.slice(sumulas.search('INSTALAÇÃO') + 11,sumulas.search('torna público')).slice(0)
      nomes = current.slice(0,current.search('CNPJ'))
    }

    if (sumulas.match('OPERAÇÃO')) {
      const current =  sumulas.slice(sumulas.search('OPERAÇÃO') + 9,sumulas.search('torna público'))
      nomes = current.slice(0,current.search('CNPJ'))
    }

    if (sumulas.match('PRÉVIA')) {
      const current = sumulas.slice(sumulas.search('PRÉVIA') + 7,sumulas.search('torna público'))
      nomes = current.slice(0,current.search('CNPJ'))
    }

    if (sumulas.match('SIMPLIFICADA')) {
      const current = sumulas.slice(sumulas.search('SIMPLIFICADA') + 13,sumulas.search('torna público'))
      nomes = current.slice(0,current.search('CNPJ')) || current.slice(0,current.search('CNPJ'))
    }

    worksheet.cell(2 + index,1).string(protocolos)
    worksheet.cell(2 + index,2).string(cnpj)
    worksheet.cell(2 + index,4).string(nomes)
    worksheet.cell(2 + index,5).string(sumulas)
    worksheet.cell(2 + index,3).string(edit)
  })


  


  /*
  rows.map((sumulas,index) => {
    const edicao = data.toString().match(/Edição\snº\s\d{5}/g)
    //worksheet.cell(2,3).string(edicao)

    data.map((cell,index) => {
      const sumulas = cell
      console.log(sumulas)
    })

    console.log(edicao)
  })
  */
  
  /*
  const diarios = [];

  rows.map(data => {
    diarios.push(data[11])
  })

  //rows.map(row => diarios.push(row[11]))
  
  const sumulas = [...new Set(diarios)].join().toString().match(/SÚMULA(.*?)\d{5}\/2020/g);

  const protocolos = sumulas?.join('').match(/\d{5}\/2020/g)
  console.log(protocolos)

  worksheet.cell(1,1).string('Protocolos')
  worksheet.cell(1,2).string('CNPJ')
  worksheet.cell(1,3).string('Nomes')
  worksheet.cell(1,4).string('Súmulas')

  sumulas?.map((data,index) => {
    const self = {
      protocolos:'',
      CNPJ:'',
      nome:'',
    }

    self.protocolos = data.match(/\d{5}\/2020/g)[0]

    const CNPJ = data.match(/[0-9]{2}\.?[0-9]{3}\.?[0-9]{3}\/?[0-9]{4}\-?[0-9]{2}/g) as []

    if (CNPJ != null) {
      self.CNPJ = CNPJ
    }

    if (data.match('INSTALAÇÃO')) {
      const current = data.slice(data.search('INSTALAÇÃO') + 11,data.search('torna público')).slice(0)
      self.nome = current.slice(0,current.search('CNPJ'))
    } else 

    if (data.match('OPERAÇÃO')) {
      const current =  data.slice(data.search('OPERAÇÃO') + 9,data.search('torna público'))
      self.nome = current.slice(0,current.search('CNPJ'))
    }


    if (data.match('PRÉVIA')) {
      const current = data.slice(data.search('PRÉVIA') + 7,data.search('torna público'))
      self.nome = current.slice(0,current.search('CNPJ'))
    }

    if (data.match('SIMPLIFICADA')) {
      const current = data.slice(data.search('SIMPLIFICADA') + 13,data.search('torna público'))
      self.nome = current.slice(0,current.search('CNPJ'))
    }

    worksheet.cell(2 + index,1).string(self.protocolos)
    worksheet.cell(2 + index,2).string(self.CNPJ)
    worksheet.cell(2 + index,3).string(self.nome)
    worksheet.cell(2 + index,4).string(data)

    console.log(self)
  })

  /*
  // protoculos
  worksheet.cell(1,1).string('Protocolo')
  protocolos?.map((data,index) => {
    worksheet.cell(2 + index,1).string(data)
  })


  // nomes

  worksheet.cell(1,4).string('nomes')
  sumulas?.map((data,index) => {

    const self = {}



    //worksheet.cell(2 + index,4).string(data)
  })

  // sumulas
  worksheet.cell(1,4).string('SUMULAS')
  sumulas?.map((data,index) => {
    worksheet.cell(2 + index,4).string(data)
  })
 


  
  */
  workbook.write('./output/Excel.xlsx');
  console.log('xlsx')
};

app();
