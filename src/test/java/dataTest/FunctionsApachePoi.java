package dataTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
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
}

