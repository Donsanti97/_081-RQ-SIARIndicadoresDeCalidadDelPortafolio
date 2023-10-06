package dataTest.historicoCarteraComercialPorOF;

import org.apache.poi.util.IOUtils;
import org.testng.annotations.Test;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;
import static dataTest.FunctionsApachePoi.runtime;

public class HistoricoCarteraComercialPorOF {
    //34 Hojas

    private static final String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";
    @Test
    public static void configuracion(){
        try {
            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            carteraBruta(excelFilePath);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 0, 0);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 1, 7);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 8, 15);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 16, 30);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 31, 60);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 61, 90);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 91, 120);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 121, 150);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 151, 180);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 181, 360);
            waitSeconds(5);

            diasDeMoraDias(excelFilePath, 361, 5000);
            waitSeconds(5);

            calificacion(excelFilePath, "A");
            waitSeconds(5);

            calificacion(excelFilePath, "B");
            waitSeconds(5);

            calificacion(excelFilePath, "C");
            waitSeconds(5);

            calificacion(excelFilePath, "D");
            waitSeconds(5);

            calificacion(excelFilePath, "E");
            waitSeconds(5);

            reEstCapital(excelFilePath);
            waitSeconds(5);

            reEstCapital(excelFilePath, 0, 150);
            waitSeconds(5);

            reEstCapital(excelFilePath, 151, 5000);
            waitSeconds(5);

            reEstNCreditos(excelFilePath);
            waitSeconds(5);

            nCreditosVigentes(excelFilePath);
            waitSeconds(5);

            clientesComercial(excelFilePath);
            waitSeconds(5);

            colocacionComercial(excelFilePath);
            waitSeconds(5);

            nCreditoComercial(excelFilePath);
            waitSeconds(5);

            colocacionPromComercial(excelFilePath);
            waitSeconds(5);

            /*comercialPercentil05(excelFilePath);
            waitSeconds(5);*/

            /*comercialPercentil08(excelFilePath);
            waitSeconds(5);*/

            comercialPzoProm(excelFilePath);
            waitSeconds(5);

            /*comercialPzoPerc05(excelFilePath);
            waitSeconds(5);*/

            /*comercialPzoPerc08(excelFilePath);
            waitSeconds();*/









        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    public static void carteraBruta(String filePath) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin);

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
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "suma");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void diasDeMoraDias(String filePath, int rangoDesde, int rangoHasta) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        String campoDiasDeMora = "dias_de_mora";

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            //int rangoDesde = 361;
            //int rangoHasta = 5000L;

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, campoDiasDeMora, rangoDesde, rangoHasta);

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
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "suma");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();
        }
    }

    public static void calificacion(String filePath, String calificacion) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        String campoCalificacion = "calificacion";

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, campoCalificacion, calificacion, calificacion);

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
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "suma");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void reEstCapital(String filePath) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        String reEstCapital = "re_est";
        String diasDeMora = "dias_de_mora";

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, reEstCapital, 1, 1);

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
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "suma");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void reEstCapital(String filePath, int diasMoradesde, int diasMoraHasta) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        String reEstCapital = "re_est";
        String diasDeMora = "dias_de_mora";

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin);

            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, reEstCapital, 1, 1, diasDeMora, diasMoradesde, diasMoraHasta);

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
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "suma");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void reEstNCreditos(String filePath) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "suma");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void nCreditosVigentes(String filePath) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "RECUENTO");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void clientesComercial(String filePath) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void colocacionComercial(String filePath) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void nCreditoComercial(String filePath) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "RECUENTO");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    public static void colocacionPromComercial(String filePath) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "PROMEDIO");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }


    //Mertodos a los que hay que hacerle un método aparte en la tabla dinámica para hallar el porcentaje 50%
    public static void comercialPercentil05(/*String filePath*/) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    //Mertodos a los que hay que hacerle un método aparte en la tabla dinámica para hallar el porcentaje 80%

    public static void comercialPercentil08(/*String filePath*/) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    //Pendiente campo no existe en OKCartera
    /*public static void bdCastigos(String filePath) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("centro_de_costo", "monto");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }
            runtime();

        }
    }*/

    //Pendiente campo no existe en OKCartera
    /*public static void clientesCastigados(String filePath) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("centro_de_costo", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

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
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }
            runtime();

        }
    }*/

    public static void comercialPzoProm(String filePath) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }

    //Mertodos a los que hay que hacerle un método aparte en la tabla dinámica para hallar el porcentaje 50%
    @Test
    public static void comercialPzoPerc05(/*String filePath*/) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados, 50);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("AQUÍ COMIENZA SUMA DE CAMPOS");
            System.out.println(camposDeseados.get(0) + ": " +camposDeseados.get(1));
            Map<String, String> resultado = calcularSumaPorValoresUnicos(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), 50);

            for (Map.Entry<String, String> entry : resultado.entrySet()){
                System.out.println("Valor único del primer encabezado: " + entry.getKey());
                System.out.println("Suma correspondiente: " + entry.getValue());
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

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
    //Mertodos a los que hay que hacerle un método aparte en la tabla dinámica para hallar el porcentaje 80%
    public static void comercialPzoPerc08(/*String filePath*/) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(filePath);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(filePath, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "modalidad";
            String valorInicio = "COMERCIAL"; // Reemplaza con el valor de inicio del rango
            String valorFin = "COMERCIAL"; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/03/2023";
            String fechafin = "31/03/2023";
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(filePath, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

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
        //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

        // Crear una nueva hoja Excel con los datos filtrados
        String nuevaHojaFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TemporalFile.xlsx"; // Reemplaza con la ruta y nombre de tu nuevo archivo Excel
        crearNuevaHojaExcel(nuevaHojaFilePath, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(nuevaHojaFilePath);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(nuevaHojaFilePath, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(nuevaHojaFilePath, sheetName, camposDeseados, 80);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            System.out.println("AQUÍ COMIENZA SUMA DE CAMPOS");
            System.out.println(camposDeseados.get(0) + ": " +camposDeseados.get(1));
            Map<String, String> resultado = calcularSumaPorValoresUnicos(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), 80);

            for (Map.Entry<String, String> entry : resultado.entrySet()){
                System.out.println("Valor único del primer encabezado: " + entry.getKey());
                System.out.println("Suma correspondiente: " + entry.getValue());
                System.out.println();
            }

            System.out.println("---------------------- CREACION TABLA DINAMICA CARTERA BRUTA");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");


            //Arreglar error al finalizar todos

            /*System.out.println("Analizando tablas dinamicas-----------");


            Map<String, Integer> dataTable = extractPivotTableData(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1));
            System.out.println("Se supone que son los datos de la dinámica---------------------------------------");

            for (Map.Entry<String, Integer> entry : dataTable.entrySet()){
                System.out.println("Claves:" + entry.getKey() + ", Value: " + entry.getValue());
            }*/
            runtime();

        }
    }



}
