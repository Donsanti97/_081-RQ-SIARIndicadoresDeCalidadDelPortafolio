package dataTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;
import org.apache.poi.util.*;

import java.io.*;
import java.util.*;

public class FunctionsApachePoi {


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
            e.printStackTrace();
        }
        return data;
    }

    //Metodo para creación de tablas dinámicas
    public static void tablasDinamicasApachePoi(/*String filepath, String hoja, String*/) throws IOException {
        String file1 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx";//OKCARTERA.20230426
        String sName = "Hoja1";

        try {
            IOUtils.setByteArrayMaxOverride(300000000);

            InputStream fileInputStream = new FileInputStream(file1);
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            //Definir hoja
            Sheet sheet = workbook.getSheet(sName);

            //Generar el area de los datos
            CellReference topLeft = new CellReference(sheet.getFirstRowNum(), sheet.getRow(sheet.getFirstRowNum()).getFirstCellNum());
            CellReference bottomRight = new CellReference(sheet.getLastRowNum(), sheet.getRow(sheet.getLastRowNum()).getLastCellNum() - 1);
            AreaReference source = new AreaReference(topLeft, bottomRight, sheet.getWorkbook().getSpreadsheetVersion());
            System.out.println(source);


            CellReference pivotCellReference = new CellReference(2, bottomRight.getCol() + 2);


            int startColIndex = pivotCellReference.getCol();
            int endColIndex = startColIndex + 1;
            // Eliminar las columnas donde se creó la tabla dinámica
            for (int i = endColIndex; i >= startColIndex; i--) {
                for (Row row : sheet) {
                    Cell cell = row.getCell(i);
                    if (cell != null) {
                        row.removeCell(cell);
                    }
                }
            }
            /*XSSFSheet xssfSheet = (XSSFSheet) sheet;
            for (XSSFPivotTable pivotTable : xssfSheet.getPivotTables()) {
                // Obtiene la celda de inicio de la tabla dinámica
                CellReference topLeftCell = pivotTable.getCTPivotTableDefinition().getRef();
                int firstRow = topLeftCell.getRow();
                int firstCol = topLeftCell.getCol();

                // Obtiene la celda de fin de la tabla dinámica
                int lastRow = firstRow + pivotTable.getRowCount();
                int lastCol = firstCol + pivotTable.getColCount();

                // Imprime el rango de celdas de la tabla dinámica
                System.out.println("Rango de la tabla dinámica:");
                System.out.println("Desde: " + new CellReference(firstRow, firstCol));
                System.out.println("Hasta: " + new CellReference(lastRow - 1, lastCol - 1));
            }*/

            // Buscar y eliminar cualquier tabla dinámica existente en la hoja de cálculo
            /*XSSFSheet xssfSheet = (XSSFSheet) sheet;
            List<XSSFPivotTable> pivotTables = xssfSheet.getPivotTables();
            if (!pivotTables.isEmpty()) {
                for (XSSFPivotTable pivotTable1 : pivotTables) {
                    System.out.println(pivotTable1.toString() + "----------");
                    // Eliminar las columnas donde se creó la tabla dinámica
                    for (int i = endColIndex; i >= startColIndex; i--) {
                        for (Row row : sheet) {
                            Cell cell = row.getCell(i);
                            if (cell != null) {
                                row.removeCell(cell);
                            }
                        }
                    }
                }
            }*/

            //Crea la tabla dinamica en la hoja de trabajo
            XSSFPivotTable pivotTable = ((XSSFSheet) sheet).createPivotTable(source, pivotCellReference);//DW12
            pivotTable.addRowLabel(0);//Agregar etiqueta de fila para el campo Modalidad (12)
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2, "Suma de Cantidad");//Agrega la columna de la que se va a hacer la suma y la etiqueta de la funcion suma(15)



            //Guardar excel
            FileOutputStream fileout = new FileOutputStream(file1);
            workbook.write(fileout);
            fileout.close();

            //Se cierra excel
            workbook.close();

            System.out.println("Tabla dinamica creada");


        } catch (IOException e) {
            throw new RuntimeException(e);
        }


    }

    //Metodo para obtener los nombres de las hojas existentes en el excel
    public static List<String> obtenerNombresDeHojas(String excelFilePath) {
        List<String> sheetNames = new ArrayList<>();
        try {
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
            e.printStackTrace();
        }
        return sheetNames;
    }

    //Metodo para obtener los encabezados de una hoja excel

    //Metodo para obtener los encabezados en las hojas
    public static List<String> obtenerEncabezados(String excelFilePath, String sheetName) {
        List<String> headers = new ArrayList<>();
        try {
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
            e.printStackTrace();
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
            e.printStackTrace();
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
            e.printStackTrace();
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
            e.printStackTrace();
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
            e.printStackTrace();
        }
        return datosFiltrados;
    }

    //Método para obtener valores de los encabezados de un rango específico cada uno, el primero rango String y el segundo rango double
    public static List<Map<String, String>> obtenerValoresDeEncabezados(String excelFilePath, String sheetName, String campoFiltrar1, String valorInicio1, String valorFin1, String campoFiltrar2, double valorInicio2, double valorFin2) {
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
            e.printStackTrace();
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
            e.printStackTrace();
        }
    }

    //Método que elimina un archivo excel existente
    public static void eliminarExcel(String filepath, int waitSeconds){
        File tempFile = new File(filepath);
        int seconds = waitSeconds*1000;

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

    public static List<String> encabezadosSegundoArchivo(String encabezadoBuscado, Workbook workbook2, Workbook workbook1){


        Sheet sheet1 = workbook1.getSheetAt(0);
        Sheet sheet2 = workbook2.getSheetAt(3);
        List<String> encabezados1 = obtenerEncabezados(sheet1);
        List<String> encabezados = Collections.singletonList(encabezados1.get(0));
        Iterator<Row> rowIterator = sheet2.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String valor = obtenerValorCelda(cell);
                for (String encabezado : encabezados1)
                if (encabezadoBuscado.equals(valor)) {

                }
            }
        }



    }

    public static List<String> encontrarEncabezadosSegundoArchivo(String encabezadoBuscado, Workbook workbook2) {
        List<String> encabezadosSegundoArchivo = new ArrayList<>();

        Sheet sheet2 = workbook2.getSheetAt(3); // Hoja del segundo archivo
        Iterator<Row> rowIterator = sheet2.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String valor = obtenerValorCelda(cell);
                if (encabezadoBuscado.equals(valor)) {
                    // Cuando encuentre el encabezado buscado, toma la fila como encabezados del segundo archivo
                    Iterator<Cell> headerIterator = row.cellIterator();
                    while (headerIterator.hasNext()) {
                        Cell headerCell = headerIterator.next();
                        encabezadosSegundoArchivo.add(obtenerValorCelda(headerCell));
                    }
                    return encabezadosSegundoArchivo;
                }
            }
        }

        return encabezadosSegundoArchivo;
    }

    private static boolean buscarEncabezadoEnColumna(String encabezado, int columnIndex, Workbook workbook) {
        Sheet sheet = workbook.getSheetAt(0); // Puedes especificar el índice de la hoja del segundo archivo
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(columnIndex);
            String valor = obtenerValorCelda(cell);
            if (encabezado.equals(valor)) {
                return true;
            }
        }
        return false;
    }
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

