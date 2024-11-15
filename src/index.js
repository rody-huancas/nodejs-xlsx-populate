import XlsxPopulate from 'xlsx-populate'

/********* Crear archivos excel *********/
/*
XlsxPopulate.fromBlankAsync().then((workbook) => {
  workbook.sheet(0).cell('A1').value('Hello World')
  return workbook.toFileAsync('./salida.xlsx')
})
*/
/*
async function main() {
  const workbook = await XlsxPopulate.fromBlankAsync()

  workbook.sheet(0).cell('A1').value('Nombre')
  workbook.sheet(0).cell('B1').value('Apellido')
  workbook.sheet(0).cell('C1').value('Edad')

  workbook.sheet(0).cell('A2').value('Juan')
  workbook.sheet(0).cell('B2').value('Perez')
  workbook.sheet(0).cell('C2').value(28)

  workbook.sheet(0).cell('A3').value('María')
  workbook.sheet(0).cell('B3').value('Gomez')
  workbook.sheet(0).cell('C3').value(21)

  workbook.toFileAsync('./excel/salida.xlsx')
}
*/

/********* Leer archivos excel *********/
/*
async function main() {
  const workbook = await XlsxPopulate.fromFileAsync('./excel/salida.xlsx')
  //   leer por celda
  //   const value = workbook.sheet('Sheet1').cell('A1').value()
  //   const value2 = workbook.sheet('Sheet1').cell('A2').value()
  //   console.log(value)
  //   console.log(value2)

  //   leer todo
  //   const value = workbook.sheet('Sheet1').usedRange().value()
  //   console.log(value)

  // leer por rango
  const value = workbook.sheet('Sheet1').range('A1:B2').value()
  console.log(value)
}
*/

/********* Crear y agregar datos *********/
/*
async function main() {
  const workbook = await XlsxPopulate.fromBlankAsync()
  workbook
    .sheet(0)
    .cell('A1')
    .value([
      [1, 2, 3],
      ['Nombre', 'Apellidos', 'Edad'],
      ['Juan', 'Perez', 25],
      ['Pedro', 'Lopez', 31],
    ])
  workbook.toFileAsync('./excel/salida2.xlsx')
}
*/

/********* Editar excel *********/
async function main() {
  const workbook = await XlsxPopulate.fromFileAsync('./excel/salida2.xlsx')

  // Cambiar el nombre a una hoja
  workbook.sheet('Sheet1').name('Hoja de prueba')

  // crear una nueva hoja
  workbook.addSheet('Hoja 2')
  workbook.addSheet('Hoja 3')

  // Eliminar una hoja
  workbook.deleteSheet('Hoja 2')

  // Salida del archivo
  // workbook.toFileAsync('./excel/salida3.xlsx')

  // colocar contraseña al archivo
  workbook.toFileAsync('./excel/salida3.xlsx', {
    password: '12346',
  })
}

main()
