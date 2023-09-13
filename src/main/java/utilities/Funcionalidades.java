package utilities;

import com.aspose.cells.*;

import java.util.*;

import org.apache.commons.text.similarity.LevenshteinDistance;

public class Funcionalidades {

    private static boolean isString(Cell cell){
        if (cell.getType() == CellValueType.IS_STRING){
            return true;
        }else {
            return false;
        }
    }
    private static boolean isNumeric(Cell cell){
        if (cell.getType() == CellValueType.IS_NUMERIC){
            return true;
        }else {
            return false;
        }
    }

    public static List<Map<String, String>> getHeaderValues(String filePath, String sheetName) throws Exception {
        List<Map<String, String>> data = new ArrayList<>();
        List<String> headers = getHeaders(filePath, sheetName);
        try {
            Workbook workbook = new Workbook(filePath);
            Worksheet worksheet = workbook.getWorksheets().get(sheetName);
            int rows = worksheet.getCells().getMaxDataRow();
            //Row row = worksheet.getCells().getRows().get(rows);
            for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
                Row row = worksheet.getCells().getRow(rowIndex);
                Map<String, String> rowData = new HashMap<>();
                for (int col = 0; col <= row.getLastCell().getColumn(); col++) {
                    Cell cell = row.get(col);
                    String header = headers.get(col);
                    String value = "";
                    if (cell != null) {
                        if (isString(cell)) {
                            value = cell.getStringValue();
                        } else if (isNumeric(cell)) {
                            value = String.valueOf(cell.getDoubleValue());
                        }
                    }
                    rowData.put(header, value);
                }
                data.add(rowData);
            }
            workbook.dispose();

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return data;
    }

    public static List<String> getHeaders(String filePath, String sheetname) throws Exception {
        List<String> headers = new ArrayList<>();
        try {
            Workbook workbook = new Workbook(filePath);
            Worksheet worksheet = workbook.getWorksheets().get(sheetname);
            Row row = worksheet.getCells().getRows().get(0);

            for (int col = 0; col <= row.getLastCell().getColumn(); col++) {
                Cell cell = row.get(col);
                headers.add(cell.getStringValue());
            }
            workbook.dispose();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        return headers;
    }

    public static List<String> getWorksheets(Workbook workbook){

        List<String> sheetNames = new ArrayList<>();
        try {
            WorksheetCollection worksheets = workbook.getWorksheets();

            // Agrega los nombres de las hojas a la lista
            for (int i = 0; i < worksheets.getCount(); i++) {
                Worksheet worksheet = worksheets.get(i);
                sheetNames.add(worksheet.getName());
            }

            workbook.dispose();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return sheetNames;
    }

    //Método para hallar similitudes entre dos cadenas
    private static double calculateSimilarity(String str1, String str2, LevenshteinDistance distance) {
        int maxLen = Math.max(str1.length(), str2.length());
        return 1.0 - (double) distance.apply(str1, str2) / maxLen;
    }

    //Método que comvierte las letras en código ascii, lo ordena de menor a mayor y devuelve la cadena ordenada
    public static String convertToAsciiAndSort(String input) {
        String toLowerCaseinput = convertToLowerCase(input);
        int[] asciiValues = new int[toLowerCaseinput.length()];
        for (int i = 0; i < toLowerCaseinput.length(); i++) {
            asciiValues[i] = toLowerCaseinput.charAt(i); // Obtener el valor ASCII de cada carácter
        }

        Arrays.sort(asciiValues); // Ordenar de menor a mayor (valores ASCII)

        StringBuilder result = new StringBuilder();
        for (int value : asciiValues) {
            result.append((char) value); // Convertir el valor ASCII de nuevo a carácter
        }

        return result.toString();
    }

    //Método que convierte las letras mayúsculas en minúsuclas de un string
    public static String convertToLowerCase(String input) {
        return input.toLowerCase();
    }
}
