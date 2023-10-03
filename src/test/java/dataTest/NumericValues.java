package dataTest;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;


public class NumericValues {

    public static void runtime() {
        Runtime runtime = Runtime.getRuntime();
        long minRunningMemory = (1024 * 1024);
        if (runtime.freeMemory() < minRunningMemory) {
            System.gc();
        }
    }

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

        List<String> sheetNames = obtenerNombresDeHojas(excelFilePath);

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
            runtime();

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


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }
            runtime();

        }


    }

    @Test
    public static void deleteTempFile() {
        eliminarExcel(System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx", 5);
    }


}
