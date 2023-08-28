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


    @Test(description = "Validar lectura de archivo por filas")
    public static void lectura(){
        try{
            String rutaA = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx";
            String rutaB = System.getProperty("user.dir") + "\\documents\\finalDocument\\4. Historico Cartera COMERCIAL por OF.xlsx";

            Workbook workbookA = new Workbook(rutaA);
            WorksheetCollection collectionA = workbookA.getWorksheets();
            int worksheetIndex = 0;
            Worksheet worksheetA = collectionA.get(worksheetIndex);

            int rowsA = worksheetA.getCells().getMaxDataRow();
            int colsA = worksheetA.getCells().getMaxDataColumn();

            for (int i = 0; i <= rowsA; i++) {
                for (int j = 0; j <= colsA; j++) {

                    String sheetNameA = worksheetA.getName().replace(" ", "");
                    boolean contain = worksheetA.getName().startsWith("ee");
                    System.out.println("Worksheet: " + worksheetA.getName());
                    List<String> list = Arrays.asList(worksheetA.getCells().get(i, j).getDisplayStringValue());

                    for (int k = 0; k < list.size(); k++) {

                        Workbook workbookB = new Workbook(rutaB);
                        WorksheetCollection collectionB = workbookA.getWorksheets();
                        for (int worksheetIndex1 = 0; worksheetIndex1 < collectionB.getCount(); worksheetIndex1++) {
                            Worksheet worksheetB = collectionA.get(worksheetIndex1);

                            int rowsB = worksheetB.getCells().getMaxDataRow();
                            int colsB = worksheetB.getCells().getMaxDataColumn();

                            String sheetNameB = worksheetB.getName().replace(" ", "");




                            if (sheetNameA.equals(sheetNameB)){
                                for (int l = 0; l < rowsB; l++) {
                                    for (int m = 0; m < colsB; m++) {
                                        ArrayList<String> listB = (ArrayList<String>) Arrays.asList(worksheetB.getCells().get(l, m).getDisplayStringValue());
                                    }
                                }
                            }
                        }




                    }

                }
            }

            List<String> filas = Arrays.asList("");


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
                System.out.println("Encabezados: " + ws.getCells().get(i).getDisplayStringValue() + "|");
                System.out.println("Row: " + row.get(i).getDisplayStringValue());
                String valores = "";
                Cell cell1 = ws.getCells().get(i);
                System.out.println("Cell1: " + cell1.getValue());


                for (int j = 0; j < rows; j++) {
                    Cell cell = ws.getCells().get(i+1, j);
                    valores = ws.getCells().checkRow(j+1).get(i).getDisplayStringValue();
                    System.out.println("Cell: " + cell.getValue());
                    if (valores.isEmpty()) {
                        valores = "0";
                    }
                    if (cell.toString().isEmpty()){
                        cell.putValue("0");
                    }
                    //System.out.println("Valores: " + valores);
                    //datosExcel.put();
                    //System.out.println("Datos Excel: " + datosExcel.put(encabezados, valores));
                    //System.out.println("Datos Excel-e: " + datosExcel.get(encabezados));
                    datosExcel1.put(cell1, cell);
                    System.out.println("DatosExcel 1: " + datosExcel1.get(cell1).getValue());
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

