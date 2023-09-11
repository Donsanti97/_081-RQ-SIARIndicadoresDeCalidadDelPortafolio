package dataTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;
import org.apache.poi.util.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

public class FunctionsApachePoi {

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

    @Test
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
            pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 2, "Suma de Cantidad");//Agrega la columna de la que se va a hacer la suma y la etiqueta de la funcion suma(15)
            pivotTable.addColLabel(1);

            //pivotTable.getCTPivotTableDefinition().getFilters();

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
}

