package dataTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.*;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.*;
import java.util.*;

public class FunctionsApachePoi {


    private static final Logger logger = LogManager.getLogger(FunctionsApachePoi.class);




    //Metodo para obtener los valores de encabezados generales
    public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String header = headers.get(cellIndex);
                    String value = "";
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            value = cell.getStringCellValue();
                            break;
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            value = String.valueOf(cell.getNumericCellValue());
                        }
                    }
                    rowData.put(header, value);
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);;
        }
        return data;
    }

    public static void convertirExcel(String archivo) throws IOException {
        FileInputStream fis = new FileInputStream(archivo);
        Workbook workbook = new XSSFWorkbook(fis);

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        try {
                            double valorNumerico = Double.parseDouble(cell.getStringCellValue());
                            // Si se puede convertir a número, establece el valor numérico
                            cell.setCellValue(valorNumerico);
                        } catch (NumberFormatException e) {
                            // No se pudo convertir a número, no hacemos nada
                        }
                    }
                }
            }
        }

        fis.close();

        // Guardar el archivo Excel con los valores convertidos
        FileOutputStream fos = new FileOutputStream(archivo);
        workbook.write(fos);
        fos.close();

        workbook.close();
    }

    //@Test
    //Metodo para creación de tablas dinámicas
    public static void tablasDinamicasApachePoi(String filePath, String codSucursal, String colValores) throws IOException {
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx";//OKCARTERA.20230426

        try {
            IOUtils.setByteArrayMaxOverride(300000000);

            convertirExcel(filePath);

            InputStream fileInputStream = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            //Definir hoja
            Sheet sheet = workbook.getSheetAt(0);

            //String codSucursal = "producto";//codigo_sucursal
            //String colValores = "cantidad";//capital

            List<String> headers = obtenerEncabezados(filePath, sheet.getSheetName());
            int index = 0;
            int index2 = 0;
            for (int i = 0; i < headers.size(); i++) {
                String header = headers.get(i);
                if(header.contains(codSucursal)){
                    index = i;
                    System.out.println("Index1: " + index);
                }
            }
            for (int i = 0; i < headers.size(); i++) {
                String header = headers.get(i);
                if(header.contains(colValores)){
                    index2 = i;
                    System.out.println("Index2: " + index2);

                }
            }


            //Generar el area de los datos
            CellReference topLeft = new CellReference(sheet.getFirstRowNum(), sheet.getRow(sheet.getFirstRowNum()).getFirstCellNum());
            CellReference bottomRight = new CellReference(sheet.getLastRowNum(), sheet.getRow(sheet.getLastRowNum()).getLastCellNum() - 1);
            AreaReference source = new AreaReference(topLeft, bottomRight, sheet.getWorkbook().getSpreadsheetVersion());
            System.out.println(source);


            CellReference pivotCellReference = new CellReference(2, bottomRight.getCol() + 2);

            //Crea la tabla dinamica en la hoja de trabajo
            XSSFPivotTable pivotTable = ((XSSFSheet) sheet).createPivotTable(source, pivotCellReference);//DW12
            pivotTable.addRowLabel(index);//Agregar etiqueta de fila para el campo Modalidad (12)
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, index2, "Suma de " + colValores);//Agrega la columna de la que se va a hacer la suma y la etiqueta de la funcion suma(15)



            //Guardar excel
            FileOutputStream fileout = new FileOutputStream(filePath);
            workbook.write(fileout);
            fileInputStream.close();
            fileout.close();


            //Se cierra excel
            workbook.close();


            System.out.println("Tabla dinamica creada");

        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
    }

    public static Map<String, Integer> extractPivotTableData(String filePath, String filterColumnName, String valueColumnName) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);
        List<XSSFTable> tables = ((XSSFSheet) sheet).getTables();

        if (tables.isEmpty()) {
            throw new IllegalArgumentException("No se encontraron tablas dinámicas en la hoja de trabajo.");
        }

        XSSFTable pivotTable = tables.get(0);
        CellReference startCell = pivotTable.getStartCellReference();
        CellReference endCell = pivotTable.getEndCellReference();
        int firstRow = startCell.getRow();
        int lastRow = endCell.getRow();

        Map<String, Integer> dataMap = new HashMap<>();

        for (int rowNum = firstRow + 1; rowNum <= lastRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            Cell filterCell = row.getCell(pivotTable.findColumnIndex(filterColumnName));
            String filterValue = filterCell.getStringCellValue();
            Cell valueCell = row.getCell(pivotTable.findColumnIndex(valueColumnName));
            int sumValue = (int) valueCell.getNumericCellValue();
            dataMap.put(filterValue, sumValue);
        }

        fis.close();
        return dataMap;
    }

    public static Map<String, Integer> processExcelFile(String filePath) throws IOException {
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0); // Suponiendo que estás trabajando en la primera hoja del archivo

        Map<String, Integer> resultMap = new HashMap<>();

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // Saltar la primera fila (encabezados)

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell tipoProductoCell = row.getCell(0); // Suponiendo que la columna 0 contiene el tipo de producto
            Cell costoCell = row.getCell(1); // Suponiendo que la columna 1 contiene el costo por producto

            String tipoProducto = tipoProductoCell.getStringCellValue();
            int costo = (int) costoCell.getNumericCellValue();

            // Verificar si ya existe la entrada en el Map
            if (resultMap.containsKey(tipoProducto)) {
                // Si existe, agregar el costo al valor existente
                int sumaCosto = resultMap.get(tipoProducto) + costo;
                resultMap.put(tipoProducto, sumaCosto);
            } else {
                // Si no existe, agregar una nueva entrada en el Map
                resultMap.put(tipoProducto, costo);
            }
        }

        fis.close();
        return resultMap;
    }

    //Metodo para obtener los nombres de las hojas existentes en el excel
    public static List<String> obtenerNombresDeHojas(String excelFilePath) {
        List<String> sheetNames = new ArrayList<>();
        try {
            IOUtils.setByteArrayMaxOverride(300000000);

            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                sheetNames.add(sheet.getSheetName());
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return sheetNames;
    }


    //Metodo para obtener los encabezados en las hojas
    public static List<String> obtenerEncabezados(String excelFilePath, String sheetName) {
        List<String> headers = new ArrayList<>();
        try {
            IOUtils.setByteArrayMaxOverride(300000000);
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            Row headerRow = sheet.getRow(0);
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return headers;
    }

    //Metodo para obtener los valores de encabezados específicos
    public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, List<String> camposDeseados) {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                    Cell cell = row.getCell(cellIndex);
                    String header = headers.get(cellIndex);
                    String value = "";
                    if (cell != null) {
                        if (cell.getCellType() == CellType.STRING) {
                            value = cell.getStringCellValue();
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            value = String.valueOf(cell.getNumericCellValue());
                        }
                    }
                    if (camposDeseados.contains(header)) {
                        rowData.put(header, value);
                    }
                }
                data.add(rowData);
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return data;
    }

    //Metodo para obtener valores de los encabezados en un rago especifico de valores
    public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar, String valorInicio, String valorFin) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex = headers.indexOf(campoFiltrar);
            if (campoFiltrarIndex == -1) {
                System.err.println("El campo especificado para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell = row.getCell(campoFiltrarIndex);
                String valorCelda = (cell != null && cell.getCellType() == CellType.STRING) ? cell.getStringCellValue() : "";
                if (valorCelda.compareTo(valorInicio) >= 0 && valorCelda.compareTo(valorFin) <= 0) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(dataCell.getNumericCellValue());
                            }
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }

    //Método para obtener valores de dos encabezados de un rango específico de valores cada uno
    public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, String valorInicio1, String valorFin1, String campoFiltrar2, String valorInicio2, String valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);
                String valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.STRING) ? cell1.getStringCellValue() : "";
                String valorCelda2 = (cell2 != null && cell2.getCellType() == CellType.STRING) ? cell2.getStringCellValue() : "";
                if (valorCelda1.compareTo(valorInicio1) >= 0 && valorCelda1.compareTo(valorFin1) <= 0 &&
                        valorCelda2.compareTo(valorInicio2) >= 0 && valorCelda2.compareTo(valorFin2) <= 0) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(dataCell.getNumericCellValue());
                            }
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }

    //Método para obtener valores de dos encabezados de un rango específico cada uno, en campos numéricos
    public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, double valorInicio1, double valorFin1, String campoFiltrar2, double valorInicio2, double valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            List<String> headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);
                double valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.NUMERIC) ? cell1.getNumericCellValue() : 0.0;
                double valorCelda2 = (cell2 != null && cell2.getCellType() == CellType.NUMERIC) ? cell2.getNumericCellValue() : 0.0;
                if (valorCelda1 >= valorInicio1 && valorCelda1 <= valorFin1 &&
                        valorCelda2 >= valorInicio2 && valorCelda2 <= valorFin2) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(dataCell.getNumericCellValue());
                            }
                        }
                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }

    //Método para obtener valores de los encabezados de un rango específico cada uno, el primero rango String y el segundo rango double
    public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, String valorInicio1, String valorFin1, String campoFiltrar2, int valorInicio2, int valorFin2) {
        List<Map<String, String>> datosFiltrados = new ArrayList<>();
        List<String> headers = null;
        try {
            IOUtils.setByteArrayMaxOverride(300000000);

            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheet(sheetName);
            headers = obtenerEncabezados(excelFilePath, sheetName);
            int campoFiltrarIndex1 = headers.indexOf(campoFiltrar1);
            int campoFiltrarIndex2 = headers.indexOf(campoFiltrar2);
            if (campoFiltrarIndex1 == -1 || campoFiltrarIndex2 == -1) {
                System.err.println("Alguno de los campos especificados para el filtro no existe.");
                return datosFiltrados;
            }

            int numberOfRows = sheet.getPhysicalNumberOfRows();
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                Cell cell1 = row.getCell(campoFiltrarIndex1);
                Cell cell2 = row.getCell(campoFiltrarIndex2);
                String valorCelda1 = (cell1 != null && cell1.getCellType() == CellType.STRING) ? cell1.getStringCellValue() : "";
                double valorCelda2 = (cell2 != null && cell2.getCellType() == CellType.NUMERIC) ? cell2.getNumericCellValue() : 0.0;
                if (valorCelda1.compareTo(valorInicio1) >= 0 && valorCelda1.compareTo(valorFin1) <= 0 &&
                        valorCelda2 >= valorInicio2 && valorCelda2 <= valorFin2) {
                    Map<String, String> rowData = new HashMap<>();
                    for (int cellIndex = 0; cellIndex < headers.size(); cellIndex++) {
                        Cell dataCell = row.getCell(cellIndex);
                        String header = headers.get(cellIndex);
                        String value = "";
                        if (dataCell != null) {
                            if (dataCell.getCellType() == CellType.STRING) {
                                value = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                value = String.valueOf(dataCell.getNumericCellValue());
                            }
                        }

                        rowData.put(header, value);
                    }
                    datosFiltrados.add(rowData);
                }
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return datosFiltrados;
    }

    //Metodo que crea una nueva hoja excel con información específica ya tratada en un archivo excel nuevo
    public static void crearNuevaHojaExcel(String filePath, List<String> headers, List<Map<String, String>> data) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NuevaHoja");

        // Crear la fila de encabezados en la nueva hoja
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
        }

        // Llenar la nueva hoja con los datos filtrados
        for (int i = 0; i < data.size(); i++) {
            Map<String, String> rowData = data.get(i);
            Row row = sheet.createRow(i + 1);
            for (int j = 0; j < headers.size(); j++) {
                String header = headers.get(j);
                String value = rowData.get(header);
                Cell cell = row.createCell(j);
                cell.setCellValue(value);
            }
        }


        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            System.out.println("Nueva hoja Excel creada o reemplazada en: " + filePath);
            fos.close();
            workbook.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
    }

    //Método que elimina un archivo excel existente
    public static void eliminarExcel(String filepath, int waitSeconds) {
        File tempFile = new File(filepath);
        int seconds = waitSeconds * 1000;

        if (tempFile.exists()) {
            try {
                // Espera durante el tiempo especificado antes de eliminar el archivo
                Thread.sleep(seconds);

                if (tempFile.delete()) {
                    System.out.println("Archivo Excel temporal eliminado con éxito.");
                } else {
                    System.out.println("No se pudo eliminar el archivo Excel temporal.");
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                System.err.println("Error al esperar antes de eliminar el archivo temporal: " + e.getMessage());
            }
        } else {
            System.out.println("El archivo Excel temporal no existe.");
        }


    }

    /*-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
    public static List<String> obtenerEncabezados(Sheet sheet) {
        List<String> encabezados = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Aquí puedes especificar en qué fila esperas encontrar los encabezados
            // Por ejemplo, si están en la tercera fila (fila índice 2), puedes usar:
            if (row.getRowNum() == 0) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    encabezados.add(obtenerValorCelda(cell));
                }
                break; // Terminamos de buscar encabezados una vez que los encontramos
            }
        }

        return encabezados;
    }

    public static List<String> obtenerNombresDeHojas(String excelFilePath, int indexFrom) {
        List<String> sheetNames = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();
            for (int i = indexFrom; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                sheetNames.add(sheet.getSheetName());
            }
            workbook.close();
            fis.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
        return sheetNames;
    }

    public static List<String> obtenerEncabezados(Sheet sheet, int index) {
        List<String> encabezados = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Aquí puedes especificar en qué fila esperas encontrar los encabezados
            // Por ejemplo, si están en la tercera fila (fila índice 2), puedes usar:
            if (row.getRowNum() == index) {
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    encabezados.add(obtenerValorCelda(cell));
                }
                break; // Terminamos de buscar encabezados una vez que los encontramos
            }
        }

        return encabezados;
    }

    public static List<String> encontrarEncabezadosSegundoArchivo(Sheet sheet, Workbook workbook2) {
        List<String> encabezadosSegundoArchivo = new ArrayList<>();

        // Busca el primer encabezado del primer archivo en la misma columna en el segundo archivo
        for (int columnIndex = 0; columnIndex < sheet.getRow(0).getLastCellNum(); columnIndex++) {
            String primerEncabezado = obtenerValorCelda(sheet.getRow(0).getCell(columnIndex));
            if (buscarEncabezadoEnColumna(primerEncabezado, columnIndex, workbook2)) {
                Sheet segundoSheet = workbook2.getSheetAt(3); // Puedes especificar el índice de la hoja del segundo archivo
                Iterator<Row> rowIterator = segundoSheet.iterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Cell cell = row.getCell(columnIndex);
                    encabezadosSegundoArchivo.add(obtenerValorCelda(cell));
                }
                break; // Terminamos de buscar encabezados en el segundo archivo
            }
        }

        return encabezadosSegundoArchivo;
    }

    private static boolean buscarEncabezadoEnColumna(String encabezado, int columnIndex, Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(3); // Puedes especificar el índice de la hoja del segundo archivo
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnIndex);
            String valor = obtenerValorCelda(cell);
            if (!valor.equals(null) || !valor.isEmpty()) {
                valor = "0";
            }
            if (encabezado.equals(valor)) {
                return true;
            }
        }
        return false;
    }

    /*-----------------------------------------------------------------------------------------*/
    public static List<String> buscarValorEnColumna(Sheet sheet, int columnaBuscada, String valorBuscado) {
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnaBuscada);
            String valorCelda = obtenerValorCelda(cell);

            if (valorBuscado.equals(valorCelda)) {
                return obtenerValoresFila(row);
            }
        }

        return null; // Valor no encontrado en la columna especificada
    }

    public static Map<String, List<String>> obtenerValoresPorEncabezado(Sheet sheet, List<String> encabezados) {
        Map<String, List<String>> valoresPorEncabezado = new HashMap<>();

        for (String encabezado : encabezados) {
            valoresPorEncabezado.put(encabezado, new ArrayList<>());
        }

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            for (int i = 0; i < encabezados.size() && i < valoresFila.size(); i++) {
                String encabezado = encabezados.get(i);
                String valor = valoresFila.get(i);

                if (valoresPorEncabezado.containsKey(encabezado)) {
                    valoresPorEncabezado.get(encabezado).add(valor);
                }
            }
        }

        return valoresPorEncabezado;
    }


    public static List<String> obtenerValoresFila(Row row) {
        List<String> valoresFila = new ArrayList<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            valoresFila.add(obtenerValorCelda(cell));
        }
        return valoresFila;
    }
    /*-----------------------------------------------------------------------------------------------*/


    private static String obtenerValorCelda(Cell cell) {
        String valor = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    valor = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        valor = cell.getDateCellValue().toString();
                    } else {
                        valor = Double.toString(cell.getNumericCellValue());
                    }
                    break;
                case BOOLEAN:
                    valor = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    valor = cell.getCellFormula();
                    break;
                default:
                    break;
            }
        }
        return valor;
    }
}

