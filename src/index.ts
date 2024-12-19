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
              returnData = { data: [] };

              if (_isColumnsObjects) {
                const columns: string[] = [];
                for (let j = startCol.charCodeAt(0); j <= endCol.charCodeAt(0); j++) {
                  const cell = sheet.getCell(String.fromCharCode(j) + startRow);
                  columns.push(cell.value as string);
                }

                const data: { [key: string]: any[] } = {};
                for (let i = parseInt(startRow) + 1; i <= parseInt(endRow); i++) {
                  for (let j = startCol.charCodeAt(0); j <= endCol.charCodeAt(0); j++) {
                    const cell = sheet.getCell(String.fromCharCode(j) + i);
                    const columnName = columns[j - startCol.charCodeAt(0)];
                    if (!data[columnName]) {
                      data[columnName] = [];
                    }
                    data[columnName].push(cell.value);
                  }
                }
                returnData.columns = columns;
                // @ts-ignore
                returnData.data = data;
              } else {
                const data: any[] = [];
                for (let i = parseInt(startRow); i <= parseInt(endRow); i++) {
                  const row: any[] = [];
                  for (let j = startCol.charCodeAt(0); j <= endCol.charCodeAt(0); j++) {
                    const cell = sheet.getCell(String.fromCharCode(j) + i);
                    row.push(cell.value);
                  }
                  data.push(row);
                }
                returnData.data = data;
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