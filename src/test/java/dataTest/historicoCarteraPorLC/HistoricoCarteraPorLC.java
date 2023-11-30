package dataTest.historicoCarteraPorLC;

import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;
import static dataTest.MethotsAzureMasterFiles.*;

public class HistoricoCarteraPorLC {
    //44 Hojas


    public static void configuracion(String masterFile) {
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
        String azureFile = getDocument();
        /*JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        masterFile = getDocument();*/
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola el número del mes y año de corte del archivo OkCartera sin espacios (Ejemplo: 02/2023 (febrero/2023))");
        String mesAnoCorte = mostrarCuadroDeTexto();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
        String fechaCorte = mostrarCuadroDeTexto();
        JOptionPane.showMessageDialog(null, "A continuación se creará un archivo temporal " +
                "\n Se recomienda seleccionar la carpeta \"Documentos\" para esta función...");
        String tempFile = getDirectory() + "\\TemporalFile.xlsx";


        try {

            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitMinutes(5);

            carteraTotal(okCartera, masterFile, azureFile, fechaCorte, "Cartera Total", tempFile);

            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "0 Dias",0, 0, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "1 - 7 Dias",1, 7, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "8 - 15 Dias",8, 15, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "16 - 30 Dias",16, 30, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "31 - 60 Dias",31, 60, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "61 - 90 Dias",61, 90, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "91 - 120 Dias",91, 120, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "121 - 150 Dias",121, 150, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "151 - 180 Dias",151, 180, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "181 - 360 Dias",181, 360, tempFile);
            carteraDias(okCartera, masterFile, azureFile, fechaCorte, "> 361 Dias",361, 5000, tempFile);

            calificacion(okCartera, masterFile, azureFile, fechaCorte, "A", "A", "A");
            calificacion(okCartera, masterFile, azureFile, fechaCorte, "B", "B", "B");
            calificacion(okCartera, masterFile, azureFile, fechaCorte, "C", "C", "C");
            calificacion(okCartera, masterFile, azureFile, fechaCorte, "D", "D", "D");

            reEstCapital(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital", tempFile);

            reEstCapital(okCartera, 0, 30, masterFile, azureFile, fechaCorte, "Re_Est Capital < = 30", tempFile);
            reEstCapital(okCartera, 31, 5000, masterFile, azureFile, fechaCorte, "Re_Est Capítal > 31", tempFile);

            reEstNCreditos(okCartera, masterFile, azureFile, fechaCorte, "Re_Est N° Creditos", tempFile);
            nCreditosVigentes(okCartera, masterFile, azureFile, fechaCorte, "N° Creditos Vigentes", tempFile);

            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC-A", "A", tempFile);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC-B", "B", tempFile);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC-C", "C", tempFile);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC-D", "D", tempFile);
            reestructuradosCapitalLc(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC-E", "E", tempFile);

            reestructuradosCapitalLcPlazosProm(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC_Plazos_Prom", tempFile);

            reestructuradosCapitalLcPlazosMin(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC_Plazos_Min", tempFile);

            reestructuradosCapitalLcPlazosMax(okCartera, masterFile, azureFile, fechaCorte, "Re_Est Capital_LC_Plazos_Max", tempFile);

            mora1raCuotaMontoLC(okCartera, masterFile, azureFile, fechaCorte, "Mora-1raCuota_Monto_LC", tempFile);

            mora1raCuotaCantLC(okCartera, masterFile, azureFile, fechaCorte, "Mora-1raCuota_Cant_LC", tempFile);

            provisiones(okCartera, masterFile, azureFile, fechaCorte, "Provisiones", tempFile);

            clientesGeneral(okCartera, masterFile, azureFile,fechaCorte, "Clientes_General", tempFile);

            colocacion(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Colocación", tempFile);

            credPromColocacion(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Cred Prom Colocación", tempFile);

            colocacionPercentil(okCartera, masterFile, azureFile, 50, mesAnoCorte, fechaCorte, "Colocación Percentil 0.5", tempFile);

            colocacionPercentil(okCartera, masterFile, azureFile, 80, mesAnoCorte, fechaCorte, "Colocación Percentil 0.8", tempFile);

            pzoProm(okCartera, masterFile, azureFile, mesAnoCorte, fechaCorte, "Pzo_Prom", tempFile);

            pzoPercentil(okCartera, masterFile, azureFile, 50, mesAnoCorte, fechaCorte, "Pzo_Percentil 0.5", tempFile);

            pzoPercentil(okCartera, masterFile, azureFile, 80, mesAnoCorte, fechaCorte, "Pzo_Percentil 0.8", tempFile);

            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            deleteTempFile(tempFile);

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    public static void carteraTotal(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            //System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();

            System.out.println("-----------CREACION TEMPORAL-----------");
        }

        // Crear una nueva hoja Excel con los datos filtrados
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void carteraDias(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int rangIni, int rangFin, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "capital");

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            String valorInicio = "0"; // Reemplaza con el valor de inicio del rango
            String valorFin = "0"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, rangIni, rangFin);

            // Especifica los campos que deseas obtener

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + rangIni + ", " + rangFin + "]");
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void calificacion(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        String campoCalificacion = "calificacion";

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "calificacion";
            String valorInicio = "A"; // Reemplaza con el valor de inicio del rango
            String valorFin = "A"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, calificacion, calificacion);

            // Especifica los campos que deseas obtener

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + calificacion + ", " + calificacion + "]");
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void reEstCapital(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        String reEstCapital = "re_est";

        List<String> camposDeseados = Arrays.asList("linea", "capital");

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "re_est";
            String valorInicio = ""; // Reemplaza con el valor de inicio del rango
            String valorFin = ""; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, reEstCapital, 1, 1);

            // Especifica los campos que deseas obtener

            // Imprimir datos filtrados
            System.out.println("Datos filtrados por " + reEstCapital + " en el rango [" + 1 + ", " + 1 + "]");
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void reEstCapital(String okCarteraFile, int diasMoradesde, int diasMoraHasta, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        String reEstCapital = "re_est";
        String diasDeMora = "dias_de_mora";

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

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

            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, reEstCapital, 1, 1, diasDeMora, diasMoradesde, diasMoraHasta);

            // Especifica los campos que deseas obtener


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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void reEstNCreditos(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "re_est");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void nCreditosVigentes(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void reestructuradosCapitalLc(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja,String calificacion, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, "calificacion", calificacion, calificacion, "re_est", 1, 1);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void reestructuradosCapitalLcPlazosProm(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, "re_est", 1, 1);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularPromedioPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void reestructuradosCapitalLcPlazosMin(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, "re_est", 1, 1);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularMinimoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void reestructuradosCapitalLcPlazosMax(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, "re_est", 1, 1);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularMaximoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void mora1raCuotaMontoLC(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, 1, valorFin, "cuota_desde_mora", 1, 1);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void mora1raCuotaCantLC(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, 1, valorFin, "cuota_desde_mora", 1, 1);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void provisiones(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "prov_cap");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void clientesGeneral(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void colocacion(String okCarteraFile, String masterFile, String azureFile,String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }


    //Queda pendiente configuración
    /*public static void nDeCredColocacion(String okCarteraFile, String masterFile, String azureFile,String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte);


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                    }
                }

            }
            runtime();

        }
    }*/

    public static void credPromColocacion(String okCarteraFile, String masterFile, String azureFile,String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularPromedioPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void colocacionPercentil(String okCarteraFile, String masterFile, String azureFile, int percent, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), percent);

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void pzoProm(String okCarteraFile, String masterFile, String azureFile, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "plazo");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }

    public static void pzoPercentil(String okCarteraFile, String masterFile, String azureFile, int percent, String mesAnoCorte, String fechaCorte, String hoja, String tempFile) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String filePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("linea", "plazo");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "dias_de_mora";
            int valorInicio = 0; // Reemplaza con el valor de inicio del rango
            int valorFin = 5000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

            // Especifica los campos que deseas obtener
            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "re_est");

            // Imprimir datos filtrados
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
        crearNuevaHojaExcel(tempFile, camposDeseados, datosFiltrados);

        System.out.println("Analisis archivo temporal----------------------");

        sheetNames = obtenerNombresDeHojas(tempFile);

        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);

            headers = obtenerEncabezados(tempFile, sheetName);

            System.out.println("Campos disponibles " + headers);

            for (String header : headers) {
                System.out.println(header);
            }


            //List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1), percent);

            List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()) {
                for (Map<String, String> datoMF : datosMasterFile) {
                    for (Map.Entry<String, String> entry : datoMF.entrySet()) {
                        /*------------------------------------------------------------*/
                        if (entryOkCartera.getKey().contains(entry.getKey())) {

                            System.out.println("CODIGO ENCONTRADO");


                            if (!entryOkCartera.getValue().equals(entry.getValue())) {

                                System.out.println("LOS VALORES ENCONTRADOS SON DISTINTOS-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());
                            } else {

                                System.out.println("LOS VALORES ENCONTRADOS SON IGUALES-> " + entryOkCartera.getValue() + ": " + entry.getValue() + " CON RESPECTO AL CODIGO: " + entry.getKey());

                            }
                        }
                        /*-------------------------------------------------------------------*/
                    }
                }

            }
            runtime();

        }
    }


}
