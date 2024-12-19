# ExcelTables4Js

**ExcelTables4Js** es una librería que convierte tablas dentro de archivos Excel (.xlsx) en objetos JSON, permitiendo al usuario elegir entre dos formatos de salida: un arreglo de arreglos o un objeto donde las llaves son los nombres de las columnas.

Esto utiliza la librera [ExcelJS](https://www.npmjs.com/package/exceljs) para leer el archivo Excel y extraer los datos de la tabla especificada. Y simplifica el proceso de extraer objetos tipo tablas de archivos Excel.

## 🚀 Características

- Convierte una tabla específica en un archivo Excel a JSON.
- Soporta dos formatos de salida:
  1. **Array de arreglos:** Las filas son representadas como arreglos.
  2. **Objetos con llaves:** Los nombres de las columnas son las llaves y los datos de las filas son los valores.
- Fácil de usar con JavaScript o TypeScript.

## 📦 Instalación

Usa npm para instalar la librería:

```bash
npm install ExcelTables4Js
```

## 📖 Uso

```javascript
import ExcelTables4Js from 'ExcelTables4Js';


//file example e.target.files[0]
const processExcel = async (file) => {
  
  const tableName = 'MyTable'; // Nombre de la tabla dentro del archivo Excel
  const isColumnsObjects = true; // Cambiar a `false` para obtener un array de arreglos

  const result = await ExcelTables4Js(file, tableName, isColumnsObjects);

  console.log(result);
};

```
## Ejemplo de salida tipo 1
```json
{
  "data": [
    ["Header1", "Header2"],
    ["Row1Col1", "Row1Col2"],
    ["Row2Col1", "Row2Col2"]
  ]
}
```

## Ejemplo de salida tipo 2
```json
{
  "columns": ["Header1", "Header2"],
  "data": {
    "Header1": ["Row1Col1", "Row2Col1"],
    "Header2": ["Row1Col2", "Row2Col2"]
  }
}
```

## 🌟 Contribuciones

¡Las contribuciones son bienvenidas! Si tienes ideas para mejorar esta librería o encuentras algún problema, por favor sigue estos pasos:

1. **Haz un fork del proyecto** desde el repositorio oficial:  
   [MilzCode/ExcelTables4Js](https://github.com/MilzCode/ExcelTables4Js).

2. **Crea una nueva rama** para tu funcionalidad o corrección de errores:  
   ```bash
   git checkout -b nombre-de-tu-rama
  ``
3. **Haz un pull request** con tus cambios para que sean revisados.



📂 Repositorio
Encuentra el código fuente de este proyecto en GitHub:
https://github.com/MilzCode/ExcelTables4Js
