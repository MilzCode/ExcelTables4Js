import ExcelJS from 'exceljs';

/**
 * Convierte un archivo Excel a JSON.
 * 
 * @param _file - Archivo Excel a procesar.
 * @param _tableName - Nombre de la tabla dentro del Excel a extraer.
 * @param _isColumnsObjects - Si es `true`, retorna datos estructurados en objetos con columnas como claves.
 * @returns Objeto JSON con los datos extraídos o `null` si no se encuentra la tabla o hay errores.
 */
const ExcelTables4Js = async (
  _file: File, 
  _tableName: string, 
  _isColumnsObjects: boolean = false
): Promise<{columns?: string[]; data: any[]} | null> => {
  if (!_file) {
    return null;
  }

  try {
    const workbook = new ExcelJS.Workbook();
    const reader = new FileReader();

    return new Promise((resolve, reject) => {
      reader.onload = async (e: ProgressEvent<FileReader>) => {
        try {
          const arrayBuffer = e.target?.result as ArrayBuffer;
          await workbook.xlsx.load(arrayBuffer);
          const sheets = workbook.worksheets;
          let returnData: {columns?: string[]; data: any[]} | null = null;

          sheets.forEach((sheet) => {
            // @ts-ignore
            const tables = sheet.tables;
            const tableNames = Object.keys(tables);

            if (tableNames.includes(_tableName)) {
              const table = tables[_tableName].table;
              const tableRange = table.tableRef;
              const [start, end] = tableRange.split(':');
              const [startCol, startRow] = start.split(/(\d+)/);
              const [endCol, endRow] = end.split(/(\d+)/);
              const totalColumns = countColumns(startCol, endCol);
              const columnsArray = excelColumns(startCol, totalColumns);
              
    

              if (_isColumnsObjects) {
                const columns: string[] = [];
                const dataObject:any = {};
                columnsArray.forEach((column) => {
                  const cell = sheet.getCell(column + startRow);
                  const columnName = cell.value as string;
                  columns.push(columnName);
                  dataObject[columnName] = [];
                  for (let i = parseInt(startRow)+1; i <= parseInt(endRow); i++) {
                    const cell = sheet.getCell(column + i);
                    dataObject[columnName].push(cell.value);
                  }
                });
                returnData = { columns, data: dataObject };
              } else {
                //ahora como arreglo de arreglos
                const data: any[] = [];
                for (let i = parseInt(startRow); i <= parseInt(endRow); i++) {
                  // @ts-ignore
                  const row = [];
                  columnsArray.forEach((column) => {
                    const cell = sheet.getCell(column + i);
                    row.push(cell.value);
                  });
                  // @ts-ignore
                  data.push(row);
                }
                returnData = { columns: data[0], data };
              }
            }
          });
          resolve(returnData || null);
        } catch (err) {
          reject(null);
        }
      };

      reader.readAsArrayBuffer(_file);
    });
  } catch (err) {
    return null;
  }
};

export default ExcelTables4Js;

function columnToNumber(column: string): number {
  let columnNumber = 0;
  for (let i = 0; i < column.length; i++) {
    columnNumber = columnNumber * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return columnNumber;
}

function countColumns(start: string, end: string): number {
  const startNumber = columnToNumber(start);
  const endNumber = columnToNumber(end);
  return endNumber - startNumber + 1;
}

function excelColumns(startColumn: string, numberOfColumns: number): string[] {
  // Convierte la columna de inicio a un número (A=1, B=2, ..., Z=26, AA=27, ...)
  function columnToNumber(column: string): number {
    let number = 0;
    for (let i = 0; i < column.length; i++) {
      number = number * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return number;
  }

  // Convierte un número a una columna de Excel
  function numberToColumn(number: number): string {
    let column = '';
    while (number > 0) {
      number--;
      column = String.fromCharCode(number % 26 + 'A'.charCodeAt(0)) + column;
      number = Math.floor(number / 26);
    }
    return column;
  }

  const startColumnNumber = columnToNumber(startColumn);
  const result: string[] = [];

  for (let i = 0; i < numberOfColumns; i++) {
    result.push(numberToColumn(startColumnNumber + i));
  }

  return result;
}