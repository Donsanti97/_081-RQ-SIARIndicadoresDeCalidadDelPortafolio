package dataTest;

import com.aspose.cells.*;
import com.google.common.base.Splitter;
import manejador_Accion.ManejadorDataFile;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class EjercicioDeEncabezados {



    @Test
    public static void test() throws Exception {
        String filePath1 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TestData.xlsx";
        String filePath2 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TestData - copia.xlsx";

        Workbook workbook1 = new Workbook(filePath1);
        Workbook workbook2 = new Workbook(filePath2);

        Set<String> sheetNames1 = new HashSet<>();
        Set<String> duplicateSheetNames = new HashSet<>();

        for (int sheetIndex = 0; sheetIndex < workbook1.getWorksheets().getCount(); sheetIndex++) {
            String name = workbook1.getWorksheets().get(sheetIndex).getName();
            String newName = name.replaceAll("\\s", "");

            if (sheetNames1.contains(newName)) {
                duplicateSheetNames.add(newName);
            } else {
                sheetNames1.add(newName);
            }
        }

        for (int sheetIndex = 0; sheetIndex < workbook2.getWorksheets().getCount(); sheetIndex++) {
            String name = workbook2.getWorksheets().get(sheetIndex).getName();
            String newName = name.replaceAll("\\s", "");

            if (sheetNames1.contains(newName)) {
                duplicateSheetNames.add(newName);
            }
        }

        if (!duplicateSheetNames.isEmpty()) {
            System.out.println("Las siguientes hojas tienen el mismo nombre en ambos archivos:");
            for (String name : duplicateSheetNames) {
                System.out.println(name);
            }
        } else {
            System.out.println("No se encontraron hojas con el mismo nombre en ambos archivos.");
        }
    }


    public static Set<String> getHeaders(Worksheet worksheet, int headerRow) {
        Set<String> headers = new HashSet<>();
        Row row = worksheet.getCells().getRows().get(headerRow);

        for (int col = 0; col <= row.getLastCell().getColumn(); col++) {
            Cell cell = row.get(col);
                headers.add(cell.getStringValue());
        }

        return headers;
    }
    public static int findHeaderRow(Worksheet worksheet, Set<String> targetHeaders) {
        for (int row = 0; row < worksheet.getCells().getMaxDataRow() + 1; row++) {
            Set<String> headers = getHeaders(worksheet, row);

            if (headers.equals(targetHeaders)) {
                return row;
            }
        }

        throw new IllegalArgumentException("No se encontraron encabezados coincidentes en la hoja.");
    }

    @Test(description = "Validar lectura de archivo por filas")
    public static void encabezados(){
        try{
            String rutaA = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx";
            String rutaB = System.getProperty("user.dir") + "\\documents\\finalDocument\\4. Historico Cartera COMERCIAL por OF.xlsx";

            Workbook workbookA = new Workbook(rutaA);
            Workbook workbookB = new Workbook(rutaB);
            WorksheetCollection collectionA = workbookA.getWorksheets();
            int worksheetIndex = 0;
            //Worksheet worksheetA = collectionA.get(worksheetIndex);
            Worksheet worksheetA = workbookA.getWorksheets().get(worksheetIndex);
            Worksheet worksheetB = workbookB.getWorksheets().get(3);
            Set<String> headers1 = new HashSet<>(getHeaders(worksheetA, 0));
            String s = headers1.toString();

            Set<String> headers2 = new HashSet<>(getHeaders(worksheetB, 168/*findHeaderRow(worksheetB, headers1)*/));
            String se1 = headers2.toString();
            switch (se1){
                case "ene":
                    se1 = s.replaceAll("ene", "01");
                    break;
                case "feb":
                    se1 = s.replaceAll("feb", "02");
                    break;
                case "mar":
                    se1 = s.replaceAll("mar", "03");
                    break;
                case "abr":
                    se1 = s.replaceAll("abr", "04");
                    break;
                case "may":
                    se1 = s.replaceAll("may", "05");
                    break;
                case "jun":
                    se1 = s.replaceAll("jun", "06");
                    break;
                case "jul":
                    se1 = s.replaceAll("jul", "07");
                    break;
                case "ago":
                    se1 = s.replaceAll("ago", "08");
                    break;
                case "sep":
                    se1 = s.replaceAll("sep", "09");
                    break;
                case "oct":
                    se1 = s.replaceAll("oct", "10");
                    break;
                case "nov":
                    se1 = s.replaceAll("nov", "11");
                    break;
                case "dic":
                    se1 = s.replaceAll("dic", "12");
                    break;
            }
            System.out.println("Encabezados 1: " + headers1);
            System.out.println(se1);
            System.out.println();
            System.out.println("Encabezados 2; " + headers2);


            //Set<String> commonHeaders = new HashSet<>(headers1);
            //commonHeaders.retainAll(headers2);

            /*if (!commonHeaders.isEmpty()) {
                System.out.println("Los siguientes encabezados se1 encuentran en ambas hojas:");
                for (String header : commonHeaders) {
                    System.out.println(header);
                }
            } else {
                System.out.println("No se1 encontraron encabezados comunes en ambas hojas.");
            }*/

            System.out.println("------------------------------------------------------------------------------");

            int rowsA = worksheetA.getCells().getMaxDataRow();
            int colsA = worksheetA.getCells().getMaxDataColumn();

            String encabezados = "";
            String valores = "";

            List<String> array = Collections.singletonList("");
            Cells cells = worksheetA.getCells();

            Map<String, String> datos = new HashMap<>();


            for (int i = 0; i < rowsA; i++) {
                for (int j = 0; j < colsA; j++) {
                    //encabezados = cells.get(0, j).getStringValue();
                    //System.out.print(encabezados + "||");

                }
                break;
                //encabezados = cells.get(i).getStringValue(); worksheetA.getCells().get(i).getStringValue();
            }


            for (int i = 0; i < rowsA; i++) {
                for (int j = 0; j < colsA; j++) {
                    valores = worksheetA.getCells().get(i+1, j).getDisplayStringValue();
                    //System.out.println("Valores: " + valores);
                    if (valores.isEmpty()){
                        valores = "0";
                    }
                    datos.put(encabezados, valores);
                    System.out.print(datos.get(encabezados) + "||");
                }
                System.out.println();


            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    @Test(description = "Validar titulos")
    public static void validaTitulos(){
        try {
            String ruta = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TestData.xlsx";
            Map<String, String> datosExcel = new HashMap<>();
            Map<Cell, Cell> datosExcel1 = new HashMap<>();
            Workbook wb = new Workbook(ruta);
            WorksheetCollection collection = wb.getWorksheets();;
            Worksheet ws = collection.get("CER150");
            int rows = ws.getCells().getMaxDataRow();
            int cols = ws.getCells().getMaxDataColumn();
            for (int i = 0; i <= cols; i++) {

                Row row = ws.getCells().checkRow(i);
                String encabezados = ws.getCells().get(i).getDisplayStringValue();
                System.out.println("Encabezados: " + encabezados/*ws.getCells().get(i).getDisplayStringValue()*/ + "|");
                //System.out.println("Row: " + row.get(i).getDisplayStringValue());
                String valores = "";
                Cell cell1 = ws.getCells().get(i);
                //System.out.println("Cell1: " + cell1.getValue());


                for (int j = 0; j < rows; j++) {
                    Cell cell = ws.getCells().get(i+1, j);
                    valores = ws.getCells().checkRow(j+1).get(i).getDisplayStringValue();
                    //System.out.println("Cell: " + cell.getValue());
                    if (valores.isEmpty()) {
                        valores = "0";
                    }
                    //if (cell.toString().isEmpty()){
                    //    cell.putValue("0");
                    //}
                    //System.out.println("Valores: " + valores);
                    datosExcel.put(encabezados, valores);
                    //System.out.println("Datos Excel: " + datosExcel.put(encabezados, valores));
                    System.out.println("Datos Excel-e: " + datosExcel.get(encabezados));
                    //datosExcel1.put(cell1, cell);
                    //System.out.println("DatosExcel 1: " + datosExcel1.get(cell1).getValue());
                }

            }
            System.out.println(datosExcel);


        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static void mapeito(){}

    @Test(description = "Matriz")
    public static void imprimirDatos() {
        try {

            int rows = 10;
            int cols = 10;

            // Crear dos matrices de 10x10
            int[][] matrix1 = new int[rows][cols];
            int[][] matrix2 = new int[rows][cols];

            // Llenar las matrices con valores aleatorios (solo para propósitos de ejemplo)
            fillMatrix(matrix1);
            fillMatrix(matrix2);

            // Comparar las filas de la misma posición en ambas matrices
            for (int i = 0; i < rows; i++) {
                for (int j = 0; j < rows - 1; j++) {
                    if (compareRows(matrix1[i], matrix2[j])) {
                        System.out.println("La fila " + i + " y la fila " + j + " son iguales en ambas matrices.");
                        System.out.println(matrix1[i] + "\n" + matrix2[i]);
                    } /*else {
                        System.out.println("La fila " + i + " y la fila " + j + " son diferentes en ambas matrices.");

                        System.out.println(matrix1[i] + "\n" + matrix2[i]);
                    }*/

                }
                /*if (compareRows(matrix1[i], matrix2[i])) {
                    System.out.println("La fila " + i + " es igual en ambas matrices.");
                    System.out.println(matrix1[i] + "\n" + matrix2[i]);
                } else {
                    System.out.println("La fila " + i + " es diferente en ambas matrices.");
                    System.out.println(matrix1[i] + "\n" + matrix2[i]);
                }*/
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    // Llenar una matriz con valores aleatorios
    public static void fillMatrix(int[][] matrix) {
        Scanner sc = new Scanner(System.in);
        int rows = matrix.length;
        int cols = matrix[0].length;

        for (int i = 0; i < rows; i++) {
            for (int j = 0; j < cols; j++) {
                //matrix[i][j] = sc.nextInt();
                matrix[i][j] = (int) (Math.random() * 2+1);
                //matrix[i][j] = matrix[i][j]+1;
                /*int a = j+1;
                matrix[i][j] = a;*/
                System.out.print((matrix[i][j]) + " ||");

            }
            System.out.println();
        }
    }

    // Comparar dos arreglos
    public static boolean compareRows(int[] row1, int[] row2) {
        if (row1.length != row2.length) {
            return false;
        }
        Arrays.sort(row1);
        System.out.println("Así se ve la vuelta ordenada: " + Arrays.toString(row1));
        Arrays.sort(row2);
        System.out.println("Así se ve la vuelta ordenada: " + Arrays.toString(row2));

        for (int i = 0; i < row1.length; i++) {
            if (row1[i] != row2[i]) {
                return false;
            }
        }

        return true;
    }
}

