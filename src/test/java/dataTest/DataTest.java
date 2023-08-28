package dataTest;

import com.aspose.cells.*;
import org.testng.annotations.Test;


import java.math.BigDecimal;
import java.util.HashSet;
import java.util.Set;
import org.apache.commons.text.similarity.LevenshteinDistance;

public class DataTest {

    private static double calculateSimilarity(String str1, String str2, LevenshteinDistance distance) {
        int maxLen = Math.max(str1.length(), str2.length());
        return 1.0 - (double) distance.apply(str1, str2) / maxLen;
    }


    @Test(description = "Consulta de campos tipo String")
    public static void consultaDatos() {
        try {


            //File file = new File("4. Historico Cartera COMERCIAL por OF.xlsx");
            //File newDir = new File("procesedDocuments");

            String file1 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx";
            String file2 = System.getProperty("user.dir") + "\\documents\\finalDocument\\4. Historico Cartera COMERCIAL por OF.xlsx";

            String test1 = "4. Historico Cartera COMERCIAL por OF.xlsx";
            String test2 = "4. Historico Cartera COMERCIAL por OF.xlsx";

            System.out.println(file1 + "------");
            System.out.println(file2 + "-------");


            if (getInformation(test1).equals(getInformation(test2))) {
                System.out.println("Compatibles");
            } else {
                System.out.println("Nel pastel");
            }


            System.out.println("------------------------------------------------------------------------------");

            //Generando archivos excel a trabajar
            Workbook workbook1 = new Workbook(file1);
            Workbook workbook2 = new Workbook(file2);

            int maxWorksheets = workbook1.getWorksheets().getCount();

            //El codigo a continuación verifica si hay nombres de hojas similares. Toma en cuenta que la mayoría del nombre contenga las letras con las que quiere coincidir
            LevenshteinDistance levenshteinDistance = new LevenshteinDistance();
            Set<String> duplicateSheetNames = new HashSet<>();

            for (int sheetIndex1 = 0; sheetIndex1 < workbook1.getWorksheets().getCount(); sheetIndex1++) {
                String name1 = workbook1.getWorksheets().get(sheetIndex1).getName();

                for (int sheetIndex2 = 0; sheetIndex2 < workbook2.getWorksheets().getCount(); sheetIndex2++) {
                    String name2 = workbook2.getWorksheets().get(sheetIndex2).getName();

                    double similarity = calculateSimilarity(name1, name2, levenshteinDistance);

                    if (similarity >= 0.5) { // Cambia este valor según la similitud deseada
                        duplicateSheetNames.add(name1);
                        duplicateSheetNames.add(name2);
                    }
                }
            }

            if (!duplicateSheetNames.isEmpty()) {
                System.out.println("Las siguientes hojas tienen nombres similares en ambos archivos:");
                for (String name : duplicateSheetNames) {
                    System.out.println(name);
                }
            } else {
                System.out.println("No se encontraron hojas con nombres similares en ambos archivos.");
            }


            //Esta secció de codigo verifica si hay nombres de hojas iguales tomando en cuenta que le quita los espacios a las hojas que tengan
            /*Set<String> sheetNames1 = new HashSet<>();
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
            }*/
            //Generando index para identificar todas las hojas dentro de los archivos
            for (int worksheetIndex = 0; worksheetIndex < maxWorksheets; worksheetIndex++) {


                Worksheet worksheet1 = workbook1.getWorksheets().get(worksheetIndex);
                Worksheet worksheet2 = workbook2.getWorksheets().get(worksheetIndex);
                System.out.println("WorkSheet1: " + worksheet1.getName());
                System.out.println("Worksheet2: " + worksheet2.getName());





                Cells cells1 = worksheet1.getCells();
                Cells cells2 = worksheet2.getCells();

                int maxRows = Math.max(cells1.getMaxDataRow(), cells2.getMaxDataRow());
                int maxCols = Math.max(cells1.getMaxDataColumn(), cells2.getMaxDataColumn());

                //Comparando si los nombres de las hojas de trabajo existen para comenzar el análisis de los campos


                    for (int row = 0; row <= maxRows; row++) {
                        for (int col = 0; col <= maxCols; col++) {
                            Cell cell1 = cells1.get(row, col);
                            Cell cell2 = cells2.get(row, col);

                            if (cell1 != null && cell2 != null) {
                                //if (cell1.getStyle().isDateTime() && cell2.getStyle().isDateTime())
                                if (cell1.getType() == CellValueType.IS_STRING && cell2.getType() == CellValueType.IS_STRING) {
                                    String value1 = cell1.getStringValue();
                                    String value2 = cell2.getStringValue();
                                    if (!value1.equals(value2)) {
                                        System.out.println("Diferencia en la fila " + (row + 1) + ", columna " + (col + 1));
                                    }
                                } else if (cell1.getType() == CellValueType.IS_NUMERIC && cell2.getType() == CellValueType.IS_NUMERIC) {
                                    double value1 = cell1.getDoubleValue();
                                    double value2 = cell2.getDoubleValue();
                                    BigDecimal decimalValue1 = new BigDecimal(value1);
                                    decimalValue1 = decimalValue1.setScale(2, BigDecimal.ROUND_HALF_UP);
                                    double doubleNum = decimalValue1.doubleValue();
                                    if (Math.abs(value1 - doubleNum) > 0.001) {
                                        System.out.println("Diferencia en la fila " + (row + 1) + ", columna " + (col + 1) + "(Numeros decimales)");
                                    }
                                    if (value1 != value2) {
                                        System.out.println("Diferencia en la fila " + (row + 1) + ", columna " + (col + 1));
                                    }
                                }

                            }

                            //System.out.print("Cell1: " + cell1.getValue()/*getDisplayStringValue()*/ + " | ");
                            //System.out.print("Cell2: " + cell2.getDisplayStringValue() + "|\n");
                        }
                    }

            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * @Description: Este método verifica las coinsidencias de los nombres de las hojas de trabajo en dos archivos Excel
     * @param workbook1
     * @param workbook2
     * @author Mairon Martinez
     * @since 28/08/2023
    **/
    public static boolean equalName(Workbook workbook1, Workbook workbook2){
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
                return true;
            }
        } else {
            System.out.println("No se encontraron hojas con el mismo nombre en ambos archivos.");
            return false;
        }
        return true;
    }

    /**
     * @Description: Este método obtiene información de un archivo excel
     * @param fileName
     * @return String
     * @author Mairon Martinez
     * @since 28/08/2023
     *
    **/
    public static String getInformation(String fileName) throws Exception {

        String string = "Valiste berenjena";

        String file = System.getProperty("user.dir") + "\\documents\\initialDocument\\" + fileName;

        Workbook wb = new Workbook(file);

        WorksheetCollection collection = wb.getWorksheets();

        for (int worksheetIndex = 0; worksheetIndex < collection.getCount(); worksheetIndex++) {
            Worksheet worksheet = collection.get(worksheetIndex);

            System.out.println("Worksheet: " + worksheet.getName());

            int rows = worksheet.getCells().getMaxDataRow();
            int cols = worksheet.getCells().getMaxDataColumn();

            for (int i = 0; i < rows; i++) {

                for (int j = 0; j < cols; j++) {
                    if (worksheet.getCells().get(i, j).getValue() != null && worksheet.getCells().get(i, j).isNumericValue()) {
                        System.out.println(worksheet.getCells().get(176, 3).getDisplayStringValue() + "||");
                        string = worksheet.getCells().get(176, 3).getDisplayStringValue() + "||";
                        return string;
                        /*if (worksheet.getCells().get(176, 8).isNumericValue()){
                            System.out.println("Es un numero pape");
                            return string;
                        }else if (!worksheet.getCells().get(176, 8).isNumericValue()){
                            System.out.println(worksheet.getCells().get(176, 8).getDisplayStringValue() + "||");
                            string = worksheet.getCells().get(176, 8).getDisplayStringValue() + "||";
                            return string;
                        }*/
                    }
                }
                System.out.println(" ");
            }
        }


        return string;
    }


}
