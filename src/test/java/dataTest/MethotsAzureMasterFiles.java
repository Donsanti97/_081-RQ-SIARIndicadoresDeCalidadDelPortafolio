package dataTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class MethotsAzureMasterFiles {

    public static String file1 = System.getProperty("user.dir") + "\\documents\\initialDocument\\Historico Cartera Comercial.xlsx";
    public static String file2 = System.getProperty("user.dir") + "\\documents\\finalDocument\\Historico Cartera COMERCIAL por OF.xlsx";


    public static void buscarYListarArchivos(String ubicacion) throws IOException {
        Path ruta = Paths.get(ubicacion);

        if (!Files.exists(ruta)) {
            System.out.println("La ubicación no existe. Creando...");
            Files.createDirectories(ruta);
            System.out.println("Ubicación creada: " + ubicacion);
        } else {
            System.out.println("La ubicación ya existe: " + ubicacion);
            listarArchivosEnCarpeta(ruta);
        }
    }

    public static void listarArchivosEnCarpeta(Path carpeta) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(carpeta)) {
            for (Path archivo : stream) {
                if (Files.isRegularFile(archivo)) {
                    System.out.println("Archivo: " + archivo.getFileName());
                }
            }
        }
    }
    /*---------------------------------------------------------------------------------------------------------------*/

    public static List<String> getWorkSheet(String filePath){
        List<String> shetNames = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();;
            for (int index = 0; index < numberOfSheets; index++) {
                Sheet sheet = workbook.getSheetAt(index);
                shetNames.add(sheet.getSheetName());
            }
            workbook.close();
            fis.close();

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return shetNames;
    }

    public static List<String> getHeaders(String excelFilePath, String sheetName) {
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

    public static List<String> getHeadersMasterfile(String excelFilePath1, String sheetName1, String excelFilePath2, String sheetName2){
        List<String> headers1 = getHeaders(excelFilePath1, sheetName1);
        String headerFirstFile1 = headers1.get(0);
        List<String> headers2 = getHeaders(excelFilePath2, sheetName2);
        String headerSecondFile = headers2.get(0);

        if (!headerFirstFile1.equals(headerSecondFile)){

        }

        return headers2;
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



    @Test
    public static void test(){
        try {
            FileInputStream fis = new FileInputStream(file1);
            Workbook workbook = new XSSFWorkbook(fis);
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook workbook2 = new XSSFWorkbook(fis2);



        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }




}
