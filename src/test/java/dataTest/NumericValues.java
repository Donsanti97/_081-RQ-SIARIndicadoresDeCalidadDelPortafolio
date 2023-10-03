package dataTest;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.*;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;


public class NumericValues {



    @Test
    public static void carteraBruta() {
        String excelFilePathTest = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleOfTheMiddleTestData.xlsx";
        String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel

        IOUtils.setByteArrayMaxOverride(300000000);
        System.out.println("URL " + excelFilePathTest);

        try {

        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }

    @Test
    public static void findFields(/*String excelFilePath, String campo, String rangoDe, String rando hasta*/) throws IOException {

        String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        String excelFilePathTest = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(excelFilePathTest);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(excelFilePathTest, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(excelFilePathTest, sheetName, campoFiltrar, valorInicio, valorFin, "dias_de_mora", 0, 0);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("----------------------");
        }

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, headers, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "suma");

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }


    }

    @Test
    public static void deleteTempFile() {
        eliminarExcel(System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx", 5);
    }

    private static final Logger logger = LogManager.getLogger(FunctionsApachePoi.class);

    public static void crearNuevaHojaExcel(String filePath, List<String> headers, List<Map<String, String>> data) throws IOException {
        try {
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




            FileOutputStream fos = new FileOutputStream(filePath);
            workbook.write(fos);
            System.out.println("Nueva hoja Excel creada o reemplazada en: " + filePath);
            fos.close();
            workbook.close();
        } catch (IOException e) {
            logger.error("Error al procesar el archivo Excel", e);
        }
    }


}
