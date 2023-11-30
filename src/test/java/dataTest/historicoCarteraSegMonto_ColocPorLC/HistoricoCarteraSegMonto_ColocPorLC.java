package dataTest.historicoCarteraSegMonto_ColocPorLC;

import org.apache.poi.util.IOUtils;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;
import static dataTest.MethotsAzureMasterFiles.*;

public class HistoricoCarteraSegMonto_ColocPorLC {
    //110 hojas
    private static String menu(java.util.List<String> opciones) {

        JFrame frame = new JFrame("Menú de Opciones");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JComboBox<String> comboBox = new JComboBox<>(opciones.toArray(new String[0]));
        comboBox.setSelectedIndex(0);

        JButton button = new JButton("Seleccionar");

        ActionListener actionListener = new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                frame.dispose(); // Cerrar la ventana después de seleccionar una opción
            }
        };
        button.addActionListener(actionListener);

        JPanel panel = new JPanel();
        panel.add(comboBox);
        panel.add(button);

        frame.add(panel);
        frame.setSize(300, 100);
        frame.setVisible(true);

        while (frame.isVisible()) {
            // Esperar hasta que la ventana se cierre
            try {
                Thread.sleep(100);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        return comboBox.getSelectedItem().toString();
    }

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
            waitSeconds(10);
            System.out.println("Espere el proceso de análisis va a comenzar...");
            waitSeconds(5);

            System.out.println("Espere un momento el análisis puede ser demorado...");
            waitSeconds(5);

            JOptionPane.showMessageDialog(null, "Para los análisis de algunas de las hojas a continuación es necesario que" +
                    "\n Digite a continuación un tipo de calificación entre [B] y [E]");
            java.util.List<String> opciones = Arrays.asList("B", "C", "D", "E");
            String calificacion = menu(opciones);


            nuevosLineas(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Lineas", tempFile);

            nuevosMay30Lineas(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_> 30_Lineas", tempFile);

            nuevosLineasBE(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Lineas_B_E", calificacion, tempFile);

            renovadoLineas(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Lineas", tempFile);

            renovadoMay30Lineas(okCartera, masterFile, azureFile, fechaCorte, "Renovado_>30_Lineas", tempFile);

            renovadoLineasBE(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Lineas_B_E", calificacion, tempFile);

            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_'0-0.5 M", 0, 5, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_0.5-1 M", 5, 10, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_1-2 M", 10, 20, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_2-3 M", 20, 30, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_3-4 M", 30, 40, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_4-5 M", 40, 50, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_5-10 M", 50, 100, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_10-15 M", 100, 150, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_15-20 M", 150, 200, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_20-25 M", 200, 250, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_25-50 M", 250, 500, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_50-100 M", 500, 1000, tempFile);
            lMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_> 100 M", 1000, 10000, tempFile);

            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_'0-0.5 M >30", 0, 5, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_0.5-1 M >30", 5, 10, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_1-2 M >30", 10, 20, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_2-3 M >30", 20, 30, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_3-4 M >30", 30, 40, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_4-5 M >30", 40, 50, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_5-10 M >30", 50, 100, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_10-15 M >30", 100, 150, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_15-20 M >30", 150, 200, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_20-25 M >30", 200, 250, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_25-50 M >30", 250, 500, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_50-100 M >30", 500, 1000, tempFile);
            lMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_>100 M >30", 1000, 10000, tempFile);

            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_'0-0.5 M B_E", 0, 5, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_0.5-1 M B_E", 5, 10, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_1-2 M B_E", 10, 20, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_2-3 M B_E", 20, 30, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_3-4 M B_E", 30, 40, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_4-5 M B_E", 40, 50, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_5-10 M B_E", 50, 100, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_10-15 M B_E", 100, 150, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_15-20 M B_E", 150, 200, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_20-25 M B_E", 200, 250, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_25-50 M B_E", 250, 500, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_50-100 M B_E", 500, 1000, calificacion, tempFile);
            lMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "L_Monto_Coloc_> 100 M B_E", 1000, 10000, calificacion, tempFile);

            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_'0-0.5 M OF", 0, 5, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_0.5-1 M OF", 5, 10, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_1-2M M OF", 10, 20, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_2-3M M OF", 20, 30, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_3-4M M OF", 30, 40, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_4-5M M OF", 40, 50, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_5-10M M OF", 50, 100, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_10-15 M OF", 100, 150, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_15-20 M OF", 150, 200, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_20-25 M OF", 200, 250, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_25-50 M OF", 250, 500, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_50-100 M OF", 500, 1000, tempFile);
            reestLC(okCartera, masterFile, azureFile, fechaCorte, "Reest_> 100 M OF", 1000, 10000, tempFile);

            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_'0-0.5 M OF", 0, 5, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_0.5-1 M OF", 5, 10, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_1-2M M OF", 10, 20, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_2-3M M OF", 20, 30, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_3-4M M OF", 30, 40, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_4-5M M OF", 40, 50, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_5-10M M OF", 50, 100, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_10-15 M OF", 100, 150, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_15-20 M OF", 150, 200, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_20-25 M OF", 200, 250, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_25-50 M OF", 250, 500, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_50-100 M OF", 500, 1000, tempFile);
            clientesLC(okCartera, masterFile, azureFile, fechaCorte, "Clientes_> 100 M OF", 1000, 10000, tempFile);

            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_'0-0.5 M OF", 0, 5, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_0.5-1 M OF", 5, 10, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_1-2M M OF", 10, 20, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_2-3M M OF", 20, 30, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_3-4M M OF", 30, 40, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_4-5M M OF", 40, 50, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_5-10M M OF", 50, 100, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_10-15 M OF", 100, 150, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_15-20 M OF", 150, 200, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_20-25 M OF", 200, 250, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_25-50 M OF", 250, 500, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_50-100 M OF", 500, 1000, tempFile);
            operacionesLC(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_> 100 M OF", 1000, 10000, tempFile);

            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_'0-0.5 M OF", 0, 5, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_0.5-1 M OF", 5, 10, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_1-2M M OF", 10, 20, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_2-3M M OF", 20, 30, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_3-4M M OF", 30, 40, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_4-5M M OF", 40, 50, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_5-10M M OF", 50, 100, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_10-15 M OF", 100, 150, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_15-20 M OF", 150, 200, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_20-25 M OF", 200, 250, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_25-50 M OF", 250, 500, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_50-100 M OF", 500, 1000, mesAnoCorte, tempFile);
            colocacion(okCartera, masterFile, azureFile, fechaCorte, "Colocación_> 100 M OF", 1000, 10000, mesAnoCorte, tempFile);





            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            deleteTempFile(tempFile);
        } catch (HeadlessException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }



    public static void nuevosLineas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "tipo_cliente";
            String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
            String valorFin = "Nuevo"; // Reemplaza con el valor de fin del rango

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
            waitSeconds(2);

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


            //List<String> camposDeseados = Arrays.asList("linea", "capital");
            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            runtime();

        }
    }

    public static void nuevosMay30Lineas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);
            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "tipo_cliente";
            String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
            String valorFin = "Nuevo"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "dias_de_mora", 31, 5000);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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
            datosFiltrados = null;
            System.gc();
            waitSeconds(2);

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            runtime();

        }
    }

    public static void nuevosLineasBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "tipo_cliente";
            String valorInicio = "Nuevo"; // Reemplaza con el valor de inicio del rango
            String valorFin = "Nuevo"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "calificacion", calificacion, calificacion);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            runtime();

        }
    }

    public static void renovadoLineas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);
            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "tipo_cliente";
            String valorInicio = "Renovado"; // Reemplaza con el valor de inicio del rango
            String valorFin = "Renovado"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            runtime();

        }
    }

    public static void renovadoMay30Lineas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja , String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);
            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "tipo_cliente";
            String valorInicio = "Renovado"; // Reemplaza con el valor de inicio del rango
            String valorFin = "Renovado"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "dias_de_mora", 31, 5000);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            runtime();

        }
    }

    public static void renovadoLineasBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "tipo_cliente";
            String valorInicio = "Renovado"; // Reemplaza con el valor de inicio del rango
            String valorFin = "Renovado"; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "calificacion", calificacion, calificacion);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInicio + ", " + valorFin + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            runtime();

        }
    }

    public static void lMontoColoc(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);
            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "valor_desem";
            int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
            int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInic + ", " + valorFinal + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            runtime();

        }
    }

    public static void lMontoColocMay30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);
            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "valor_desem";
            long valorInicio = valorInic * 1000000L; // Reemplaza con el valor de inicio del rango
            long valorFin = valorFinal * 1000000L; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "dias_de_mora", 31, 5000);

            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInic + ", " + valorFinal + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    runtime();
                    waitSeconds(2);
                }

            }
            System.gc();
            waitSeconds(2);

        }
        System.gc();
        waitSeconds(2);
    }

    public static void lMontoColocBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String calificacion, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "valor_desem";
            int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
            int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, "calificacion", calificacion, calificacion, campoFiltrar, valorInicio, valorFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInic + ", " + valorFinal + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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
            datosFiltrados = null;
            System.gc();
            waitSeconds(2);

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            System.gc();
            waitSeconds(2);

        }
        System.gc();
        waitSeconds(2);
    }

    public static void reestLC(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "capital");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "valor_desem";
            int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
            int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, "re_est", 1, 1, campoFiltrar, valorInicio, valorFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInic + ", " + valorFinal + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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
            datosFiltrados = null;
            System.gc();
            waitSeconds(2);

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            System.gc();
            waitSeconds(2);

        }
        System.gc();
        waitSeconds(2);
    }

    public static void clientesLC(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "Cliente");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "valor_desem";
            int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
            int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInic + ", " + valorFinal + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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
            datosFiltrados = null;
            System.gc();
            waitSeconds(2);

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularSumaPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            System.gc();
            waitSeconds(2);

        }
        System.gc();
        waitSeconds(2);
    }

    public static void operacionesLC(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "linea");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "valor_desem";
            int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
            int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInic + ", " + valorFinal + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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
            datosFiltrados = null;
            System.gc();
            waitSeconds(2);

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

            Map<String, String> resultado = functions.calcularConteoPorValoresUnicos(tempFile, camposDeseados.get(0), camposDeseados.get(1));

            java.util.List<Map<String, String>> datosMasterFile = obtenerValoresEncabezados2(azureFile, masterFile, hoja, fechaCorte)/*getHeadersMFile(azureFile, masterFile, fechaCorte)*/;


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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            System.gc();
            waitSeconds(2);

        }
        System.gc();
        waitSeconds(2);
    }

    public static void colocacion(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String mesAnoCorte, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        java.util.List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        java.util.List<String> headers = null;
        java.util.List<Map<String, String>> datosFiltrados = null;
        java.util.List<String> camposDeseados = Arrays.asList("linea", "valor_desem");
        for (String sheetName : sheetNames) {
            System.out.println("Contenido de la hoja: " + sheetName);
            headers = obtenerEncabezados(okCarteraFile, sheetName);

            // Listar campos disponibles
            System.out.println("Campos disponibles:");
            for (String header : headers) {
                System.out.println(header);
            }
            runtime();
            waitSeconds(2);

            // Especifica el campo en el que deseas aplicar el filtro
            String campoFiltrar = "valor_desem";
            int valorInicio = valorInic * 1000000; // Reemplaza con el valor de inicio del rango
            int valorFin = valorFinal * 1000000; // Reemplaza con el valor de fin del rango
            String fechaInicio = "01/" + mesAnoCorte;
            String fechafin = "31/" + mesAnoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechafin);

            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);


            // Especifica los campos que deseas obtener


            // Imprimir datos filtrados
            System.out.println("DATOS FILTRADOS");
            System.out.println("Datos filtrados por " + campoFiltrar + " en el rango [" + valorInic + ", " + valorFinal + "]");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);

                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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
            datosFiltrados = null;
            System.gc();
            waitSeconds(2);

            datosFiltrados = obtenerValoresDeEncabezados(tempFile, sheetName, camposDeseados);


            System.out.println("VALORES DEL OK CARTERA");
            for (Map<String, String> rowData : datosFiltrados) {
                for (String campoDeseado : camposDeseados) {
                    String valorCampo = rowData.get(campoDeseado);
                    System.out.println(campoDeseado + ": " + valorCampo);
                }
                System.out.println();
            }
            runtime();
            waitSeconds(2);

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
                        }else {
                            System.err.println("Código no encontrado: " + entryOkCartera.getKey());
                        }
                        /*-------------------------------------------------------------------*/
                    }
                    waitSeconds(5);
                    runtime();
                }

            }
            System.gc();
            waitSeconds(2);

        }
        System.gc();
        waitSeconds(2);
    }


}
