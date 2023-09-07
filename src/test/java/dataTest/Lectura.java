package dataTest;

import com.aspose.cells.Workbook;
import org.testng.annotations.Test;
import utilities.Funcionalidades;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;


public class Lectura {



    @Test
    public static void test() throws Exception {
        String file1 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx";
        Workbook workbook = new Workbook(file1);
        List<String> sheetNames = Funcionalidades.getWorksheets(workbook);


        for (String sheetName : sheetNames) {
            System.out.println(sheetName);
            List<String> headers = Funcionalidades.getHeaders(file1, sheetName);
            List<Map<String, String>> data = Funcionalidades.getHeaderValues(file1, sheetName);

            System.out.println("Encabezados");
            for (String header : headers) {
                System.out.print(header + "\t");
            }
            System.out.println();

            for (Map<String, String> rowData :
                    data) {
                for (Map.Entry<String, String> entry :
                        rowData.entrySet()) {
                    System.out.println(entry.getKey() + ": " + entry.getValue());
                }
                System.out.println();
            }

            System.out.println("-----------------------------------------------------");
        }


    }





    /*@Test
    public static void lectura(){
        String archivoCsv1 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx"; // Ruta del archivo CSV 1
        String archivoCsv2 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx"; // Ruta del archivo CSV 2

        List<Map<String, String>> datosArchivo1 = leerArchivoCSV(archivoCsv1);
        List<Map<String, String>> datosArchivo2 = leerArchivoCSV(archivoCsv2);

        List<String> logComparacion = new ArrayList<>();

        for (int indice = 0; indice < datosArchivo1.size(); indice++) {
            boolean encontrado = false;
            String logMensaje = "Objeto en posición [" + indice + "] del archivo 1 ";

            for (Map<String, String> dato2 : datosArchivo2) {
                if (sonIguales(datosArchivo1.get(indice), dato2)) {
                    logComparacion.add(logMensaje + "está presente en al menos un objeto del archivo 2.");
                    encontrado = true;
                    break;
                }
            }

            if (!encontrado) {
                logComparacion.add(logMensaje + "no está presente en ningún objeto del archivo 2.");
            }
        }

        for (String mensaje : logComparacion) {
            System.out.println(mensaje);
        }
    }

    public static List<Map<String, String>> leerArchivoCSV(String archivoCsv) {
        List<Map<String, String>> objectsList = new ArrayList<>();

        try (BufferedReader br = new BufferedReader(new FileReader(archivoCsv))) {
            String line;
            String[] headerRow = br.readLine().split(",");

            while ((line = br.readLine()) != null) {
                String[] data = line.split(",");
                Map<String, String> rowData = new HashMap<>();

                for (int i = 0; i < headerRow.length - 1; i++) {
                    rowData.put(headerRow[i], data[i]);
                }

                objectsList.add(rowData);
            }
        } catch (IOException e) {
            System.err.println("No se pudo abrir el archivo CSV.");
        }

        return objectsList;
    }

    public static boolean sonIguales(Map<String, String> objeto1, Map<String, String> objeto2) {
        return objeto1.equals(objeto2);
    }


    @Test
    public static void datos(){
        String archivoExcel1 = System.getProperty("user.dir") +
                "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx"; // Ruta del archivo Excel 1
        String archivoExcel2 = System.getProperty("user.dir") +
                "\\documents\\procesedDocuments\\Historico Cartera Comercial (3).xlsx"; // Ruta del archivo Excel 2

        List<RowData> datosArchivo1 = leerArchivoExcel(archivoExcel1);
        List<RowData> datosArchivo2 = leerArchivoExcel(archivoExcel2);

        List<String> logComparacion = new ArrayList<>();

        for (int indice = 0; indice < datosArchivo1.size(); indice++) {
            boolean encontrado = false;
            String logMensaje = "Fila en posición [" + indice + datosArchivo1.get(indice).toString() + "] del archivo 1 ";

            for (RowData dato2 : datosArchivo2) {
                if (sonIguales(datosArchivo1.get(indice), dato2)) {
                    logComparacion.add(logMensaje + "está presente en al menos una fila del archivo 2.");
                    encontrado = true;
                    break;
                }
            }

            if (!encontrado) {
                logComparacion.add(logMensaje + "no está presente en ninguna fila del archivo 2.");
            }
        }

        for (String mensaje : logComparacion) {
            System.out.println(mensaje);
        }
    }

    public static List<RowData> leerArchivoExcel(String archivoExcel) {
        List<RowData> rowsList = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(archivoExcel);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Hoja en la posición 0

            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                RowData rowData = new RowData();

                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    rowData.addColumn(cell.toString());
                }

                rowsList.add(rowData);
            }
        } catch (IOException e) {
            System.err.println("No se pudo abrir el archivo Excel.");
        }

        return rowsList;
    }

    public static boolean sonIguales(RowData fila1, RowData fila2) {
        return fila1.equals(fila2);
    }
}

class RowData {
    private List<String> columns = new ArrayList<>();

    public void addColumn(String value) {
        columns.add(value);
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof RowData)) return false;
        RowData rowData = (RowData) o;
        return columns.equals(rowData.columns);
    }

    @Override
    public int hashCode() {
        return columns.hashCode();
    }*/

}
