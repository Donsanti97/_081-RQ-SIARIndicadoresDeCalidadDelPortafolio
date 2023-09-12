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
    public static void tablasDinamicasApachePoi() throws IOException {
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

            //Crea la tabla dinamica en la hoja de trabajo
            XSSFPivotTable pivotTable = ((XSSFSheet) sheet).createPivotTable(source, new CellReference("E13"));//DW12
            pivotTable.addRowLabel(0);//Agregar etiqueta de fila para el campo Modalidad (12)
            //pivotTable.addColLabel(0, "bebida");
            //pivotTable.
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2, "Suma de Cantidad");//Agrega la columna de la que se va a hacer la suma y la etiqueta de la funcion suma(15)
            //pivotTable.addColLabel(1);

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

        int counter = 0;
        if (new File(filePath).exists()){
            counter++;
            String filename = filePath + counter + "xlsx";
            filePath = filename;
        }
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            workbook.write(fos);
            System.out.println("Nueva hoja Excel creada en: " + filePath);
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
}

