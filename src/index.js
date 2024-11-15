import xlsxPopulate from 'xlsx-populate'

// Crear archivos excel
/*
xlsxPopulate.fromBlankAsync().then((workbook) => {
  workbook.sheet(0).cell('A1').value('Hello World')
  return workbook.toFileAsync('./salida.xlsx')
})
*/
async function main() {
  const workbook = await xlsxPopulate.fromBlankAsync()

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

main()
