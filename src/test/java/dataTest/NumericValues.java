package dataTest;

import org.testng.annotations.Test;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;


public class NumericValues {

    @Test
    public static void findFields() {

        String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel

        List<String> sheetNames = obtenerNombresDeHojas(excelFilePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(excelFilePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "TipoProducto";
            String valorInicio = "bebida"; // Reemplaza con el valor de inicio del rango
            String valorFin = "bebida"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(excelFilePath, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            List<String> camposDeseados = Arrays.asList("producto", "cantidad");

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

        System.out.println("----------------------");
    }

    @Test
    public static void deleteTempFile(){
        eliminarExcel(System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx", 5);
    }


}
