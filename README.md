# ExcelTables4Js

**ExcelTables4Js** es una librería que convierte tablas dentro de archivos Excel (.xlsx) en objetos JSON, permitiendo al usuario elegir entre dos formatos de salida: un arreglo de arreglos o un objeto donde las llaves son los nombres de las columnas.

Esto utiliza la librera [ExcelJS](https://www.npmjs.com/package/exceljs) para leer el archivo Excel y extraer los datos de la tabla especificada. Y simplifica el proceso de extraer objetos tipo tablas de archivos Excel.


## Parámetros de entrada

### `_file`
- **Tipo**: `File`
- **Descripción**: Archivo Excel que se desea procesar
- **Requerido**: Sí

### `_tableName`
- **Tipo**: `string`
- **Descripción**: Nombre de la tabla dentro del archivo Excel que se desea procesar
- **Requerido**: Sí

### `_isColumnsObjects`
- **Tipo**: `boolean`
- **Descripción**: Define la estructura de salida de los datos procesados
- **Valores posibles**:
  - `true`: Los datos se transforman en un objeto donde las columnas actúan como claves
  - `false`: Los datos se transforman en un arreglo de arreglos
- **Requerido**: Sí

### `customDataFunction`
- **Tipo**: `Function`
- **Descripción**: Función personalizada para transformar los datos durante el procesamiento
- **Requerido**: No
- **Por defecto**: `undefined`

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

## Ejemplo con función personalizada
```typescript
const customFunction = (value: any, column: string) => {
  if (column === 'Price') {
    return parseFloat(value) || 0;
  }
  return value;
};

ExcelTables4Js(file, tableName, true, customFunction).then((result) => {
  console.log(result);
});
```



## Objeto con columnas como claves (_isColumnsObjects = true)
```json
{
  "columns": ["Column1", "Column2", "__fileSourceName"],
  "data": {
    "Column1": ["Value1", "Value2"],
    "Column2": ["ValueA", "ValueB"],
    "__fileSourceName": ["example.xlsx", "example.xlsx"]
  }
}

```

## Arreglo de arreglos (_isColumnsObjects = false)
```json
{
  "columns": ["Column1", "Column2", "__fileSourceName"],
  "data": [
    ["Value1", "ValueA", "example.xlsx"],
    ["Value2", "ValueB", "example.xlsx"]
  ]
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
