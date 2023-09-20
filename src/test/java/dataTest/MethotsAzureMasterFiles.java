package dataTest;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.*;
import java.util.*;

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

    public static List<String> getWorkSheet(String filePath, int i) {
        List<String> shetNames = new ArrayList<>();
        try {
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(fis);
            int numberOfSheets = workbook.getNumberOfSheets();
            ;
            for (int index = i; index < numberOfSheets; index++) {
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

    public static List<String> getHeaders(Sheet sheet) {
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

    public static List<String> getHeadersMasterfile(Sheet sheet1, Sheet sheet2) throws IOException {
        List<String> headers1 = getHeaders(sheet1);
        String headerFirstFile1 = headers1.get(0);
        List<String> headers2 = getHeaders(sheet2);
        String headerSecondFile = headers2.get(0);

        if (!headerFirstFile1.equals(headerSecondFile)) {
            headers2 = FunctionsApachePoi.buscarValorEnColumna(sheet1, 0, headers1.get(0));
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

    public static String obtenerValorCelda(Cell cell) {
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
                    valor = evaluarFormula(cell);
                    break;
                default:
                    break;
            }
        }
        return valor;
    }
    public static String evaluarFormula(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);
            return cellValue.formatAsString();
        } catch (Exception e) {
            return "";
        }
    }

    public static List<Map<String, String>> obtenerValoresPorFilas(Sheet sheet, List<String> encabezados) {
        List<Map<String, String>> valoresPorFilas = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.iterator();
        // Omitir la primera fila ya que contiene los encabezados
        if (rowIterator.hasNext()) {
            rowIterator.next();
        }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            List<String> valoresFila = obtenerValoresFila(row);

            Map<String, String> fila = new HashMap<>();
            for (int i = 0; i < encabezados.size() && i < valoresFila.size(); i++) {
                String encabezado = encabezados.get(i);
                String valor = valoresFila.get(i);
                fila.put(encabezado, valor);
            }

            valoresPorFilas.add(fila);
        }

        return valoresPorFilas;
    }


    @Test
    public static void test() {
        try {
            FileInputStream fis = new FileInputStream(file1);
            Workbook workbook = new XSSFWorkbook(fis);
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook workbook2 = new XSSFWorkbook(fis2);
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(3);

            List<String> nameSheets1 = getWorkSheet(file1, 0);
            List<String> nameSheets2 = getWorkSheet(file2, 3);

            List<String> encabezados1 = null;
            List<String> encabezados2 = null;

            List<Map<String, String>> valoresEncabezados1 = null;
            List<Map<String, String>> valoresEncabezados2 = null;


            for (String sheets : nameSheets1) {
                System.out.print("SheetName: ");
                System.out.println(sheets);
                encabezados1 = getHeaders(file1, "CER150");
                //System.out.println("Headers: ");
                for (String headers : encabezados1) {
                    //System.out.print(headers + "||");
                    valoresEncabezados1 = FunctionsApachePoi.obtenerValoresDeEncabezados(file1, sheets);
                    /*valoresEncabezados1 = obtenerValoresPorFilas(sheet1, encabezados1);
                    //System.out.println(valoresEncabezados1);
                    for (Map<String, String> map : valoresEncabezados1){
                        System.out.println("Fila: ");
                        for (Map.Entry<String, String> entry : map.entrySet()){
                            System.out.println("Headers1: " + entry.getKey() + ", value: " + entry.getValue());
                        }
                    }*/
                }
            }
            System.out.println("------------------------------------------------------------------------------------------");

            valoresEncabezados1 = obtenerValoresPorFilas(sheet1, encabezados1);
            for (Map<String, String> map : valoresEncabezados1){
                System.out.println("Fila: ");
                for (Map.Entry<String, String> entry : map.entrySet()){
                    System.out.println("Headers1: " + entry.getKey() + ", value: " + entry.getValue());
                }
            }

            System.out.println("------------------------------------------------------");
            for (String sheets2 : nameSheets2) {
                System.out.println("SheetName2: " + sheets2);
                encabezados2 = getHeadersMasterfile(sheet1, sheet2);
                for (String headers : encabezados2) {
                    System.out.print("Headers2: " + headers);
                }
            }

            System.out.println("-------------------------------------------------------------------------------------");
            valoresEncabezados2 = obtenerValoresPorFilas(sheet2, encabezados2);
            for (Map<String, String> map : valoresEncabezados2){
                System.out.println("Fila2: ");
                for (Map.Entry<String, String> entry : map.entrySet()){
                    System.out.println("Headers2: " + entry.getKey() + ", Value: " + entry.getValue());
                }
            }

            System.out.println("---------------------------------------------------------------------------------------");
            for (String e1 : encabezados1){
                for (String e2: encabezados2){
                    if (e1.equals(e2)){
                        System.out.println("equals");
                    }else {
                        System.out.println("No equals");
                    }
                }
            }



        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }


}
