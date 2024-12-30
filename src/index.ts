import ExcelJS from 'exceljs';

/**
 * Convierte un archivo Excel a JSON.
 * 
 * @param _file - Archivo Excel a procesar.
 * @param _tableName - Nombre de la tabla dentro del Excel a extraer.
 * @param _isColumnsObjects - Si es `true`, retorna datos estructurados en objetos con columnas como claves.
 * @param customDataFunction - Función personalizada para procesar los datos de las celdas.
 * @returns Objeto JSON con los datos extraídos o `null` si no se encuentra la tabla o hay errores.
 */
const ExcelTables4Js = async (
  _file: File,
  _tableName: string,
  _isColumnsObjects: boolean = false,
  customDataFunction?: (data: any,columnTitle:any) => any
): Promise<{ columns?: string[]; data: any[] } | null> => {
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
          let returnData: { columns?: string[]; data: any[] } | null = null;

          sheets.forEach((sheet) => {
            // @ts-ignore
            const tables = sheet.tables;
            const tableNames = Object.keys(tables);
            const fileName = _file.name;
            if (tableNames.includes(_tableName)) {
              const table = tables[_tableName].table;
              const tableRange = table.tableRef;
              const [start, end] = tableRange.split(':');
              const [startCol, startRow] = start.split(/(\d+)/);
              const [endCol, endRow] = end.split(/(\d+)/);
              const totalColumns = countColumns(startCol, endCol);
              const columnsArray = excelColumns(startCol, totalColumns);
              const nameFileSourceColumn = "__fileSourceName"
              const nameColumnRow = "__row"




              if (_isColumnsObjects) {
                const columns: string[] = [];
                const dataObject: any = {};
                columns.push(nameFileSourceColumn)
                dataObject[nameFileSourceColumn] = [];
                columns.push(nameColumnRow)
                dataObject[nameColumnRow] = [];
                
                columnsArray.forEach((column,idx) => {
                  const cell = sheet.getCell(column + startRow);
                  const columnName = cell.value as string;
                  columns.push(columnName);
                  dataObject[columnName] = [];
                  for (let i = parseInt(startRow) + 1; i <= parseInt(endRow); i++) {
                    const cell = sheet.getCell(column + i);
                    const value = fixerData(cell.value,columnName,customDataFunction)
                    dataObject[columnName].push(value);
                    if (
                      idx === 0
                    ) {
                      dataObject[nameFileSourceColumn].push(fileName)
                      dataObject[nameColumnRow].push(i)
                    }

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
                    const columnName = sheet.getCell(column + startRow).value as string;
                    let value = cell.value;
                    if (i !== parseInt(startRow)) {
                      value = fixerData(cell.value,columnName,customDataFunction)
                    }
                    row.push(value);
                  });
                  // @ts-ignore
                  data.push(row);
                  if (i === parseInt(startRow)) {
                    data[0].push(nameFileSourceColumn)
                    data[0].push(nameColumnRow)
                  } else {
                    data[data.length - 1].push(fileName)
                    data[data.length - 1].push(i)
                  }
                }
                returnData = { columns: data[0], data: data.slice(1) };
              }
            }
          });
          resolve(returnData || null);
        } catch (err) {
          console.error(err);
          reject(null);
        }
      };

      reader.readAsArrayBuffer(_file);
    });
  } catch (err) {
    console.error(err);
    return null;
  }
};

export default ExcelTables4Js;

// Convierte la columna de inicio a un número (A=1, B=2, ..., Z=26, AA=27, ...)
export function columnToNumber(column: string): number {
  let columnNumber = 0;
  for (let i = 0; i < column.length; i++) {
    columnNumber = columnNumber * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return columnNumber;
}

export function countColumns(start: string, end: string): number {
  const startNumber = columnToNumber(start);
  const endNumber = columnToNumber(end);
  return endNumber - startNumber + 1;
}

  // Convierte un número a una columna de Excel
export function numberToColumn(number: number): string {
    let column = '';
    while (number > 0) {
      number--;
      column = String.fromCharCode(number % 26 + 'A'.charCodeAt(0)) + column;
      number = Math.floor(number / 26);
    }
    return column;
  }

function excelColumns(startColumn: string, numberOfColumns: number): string[] {
  
  const startColumnNumber = columnToNumber(startColumn);
  const result: string[] = [];

  for (let i = 0; i < numberOfColumns; i++) {
    result.push(numberToColumn(startColumnNumber + i));
  }

  return result;
}

const fixerData = (cellValue: any,column:string,customDataFunction:any) => {
  const value_string = cellValue === null ? null : String(cellValue).trim();
  let value = value_string === "" ? null : value_string;
  if (customDataFunction) {
    const outputFn = customDataFunction(value,column)
    value = outputFn
  }
  return value
}