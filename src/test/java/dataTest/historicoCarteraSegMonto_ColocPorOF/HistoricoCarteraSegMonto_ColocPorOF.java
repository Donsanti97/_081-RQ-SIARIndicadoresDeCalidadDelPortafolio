package dataTest.historicoCarteraSegMonto_ColocPorOF;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;
import static dataTest.MethotsAzureMasterFiles.*;

public class HistoricoCarteraSegMonto_ColocPorOF {
    //110 hojas

    private static String menu(List<String> opciones) {

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

            /*JOptionPane.showMessageDialog(null, "Para los análisis de algunas de las hojas a continuación es necesario que" +
                    "\n Digite a continuación un tipo de calificación entre [B] y [E]");
            List<String> opciones = Arrays.asList("B", "C", "D", "E");
            String calificacion = menu(opciones);*/

            nuevosOficinas(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas", tempFile);

            /*nuevosOficinasMay30(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas > 30", tempFile);

            nuevosOficinasBE(okCartera, masterFile, azureFile, fechaCorte, "Nuevos_Oficinas_B_E", calificacion, tempFile);

            renovadoOficinas(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas", tempFile);

            renovadoOficinasMay30(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas_>30", tempFile);

            renovadoOficinasBE(okCartera, masterFile, azureFile, fechaCorte, "Renovado_Oficinas_B_E", calificacion, tempFile);

            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc '0-0.5 M", 0, 5, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 0.5-1 M", 5, 10, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 1-2 M", 10, 20, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 2-3 M", 20, 30, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 3-4 M", 30, 40, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 4-5 M", 40, 50, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 5-10 M", 50, 100, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 10-15 M", 100, 150, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 15-20 M", 150, 200, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 20-25 M", 200, 250, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 25-50 M", 250, 500, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc 50-100 M", 500, 1000, tempFile);
            oficinasMontoColoc(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Coloc > 100 M", 1000, 10000, tempFile);

            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol '0-0.5 >30", 0, 5, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 0.5-1 > 30", 5, 10, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 1-2M >30", 10, 20, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 2-3M >30", 20, 30, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 3-4M >30", 30, 40, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 4-5M >30", 40, 50, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 5-10M >30", 50, 100, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 10-15 >30", 100, 150, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 15-20 >30", 150, 200, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 20-25 >30", 200, 250, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 25-50 >30", 250, 500, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 50-100 >30", 500, 1000, tempFile);
            oficinasMontoColocMay30(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol > 100 >30", 1000, 10000, tempFile);

            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol '0-0.5 B_E", 0, 5, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 0.5-1 B_E", 5, 10, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 1-2 B_E", 10, 20, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 2-3 B_E", 20, 30, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 3-4 B_E", 30, 40, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 4-5 B_E", 40, 50, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 5-10 B_E", 50, 100, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 10-15 B_E", 100, 150, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 15-20 B_E", 150, 200, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 20-25 B_E", 200, 250, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 25-50 B_E", 250, 500, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol 50-100 B_E", 500, 1000, calificacion, tempFile);
            oficinasMontoColocBE(okCartera, masterFile, azureFile, fechaCorte, "Oficinas_Monto_Cocol > 100 B_E", 1000, 10000, calificacion, tempFile);

            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_'0-0.5 M", 0, 5, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_0.5-1 M", 5, 10, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_1-2M M", 10, 20, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_2-3M M", 20, 30, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_3-4M M", 30, 40, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_4-5M M", 40, 50, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_5-10M M", 50, 100, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_10-15 M", 100, 150, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_15-20 M", 150, 200, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_20-25 M", 200, 250, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_25-50 M", 250, 500, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_50-100 M", 500, 1000, tempFile);
            reestOF(okCartera, masterFile, azureFile, fechaCorte, "Reest_> 100 M", 1000, 10000, tempFile);

            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_'0-0.5 M", 0, 5, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_0.5-1 M", 5, 10, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_1-2M M", 10, 20, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_2-3M M", 20, 30, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_3-4M M", 30, 40, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_4-5M M", 40, 50, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_5-10M M", 50, 100, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_10-15 M", 100, 150, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_15-20 M", 150, 200, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_20-25 M", 200, 250, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_25-50 M", 250, 500, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_50-100 M", 500, 1000, tempFile);
            clientesOF(okCartera, masterFile, azureFile, fechaCorte, "Clientes_> 100 M", 1000, 10000, tempFile);

            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_'0-0.5 M", 0, 5, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_0.5-1 M", 5, 10, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_1-2M M", 10, 20, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_2-3M M", 20, 30, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_3-4M M", 30, 40, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_4-5M M", 40, 50, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_5-10M M", 50, 100, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_10-15 M", 100, 150, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_15-20 M", 150, 200, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_20-25 M", 200, 250, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_25-50 M", 250, 500, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_50-100 M", 500, 1000, tempFile);
            operacionesOF(okCartera, masterFile, azureFile, fechaCorte, "Operaciones_> 100 M", 1000, 10000, tempFile);

            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_'0-0.5 M", 0, 5, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_0.5-1 M", 5, 10, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_1-2M M", 10, 20, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_2-3M M", 20, 30, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_3-4M M", 30, 40, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_4-5M M", 40, 50, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_5-10M M", 50, 100, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_10-15 M", 100, 150, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_15-20 M", 150, 200, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_20-25 M", 200, 250, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_25-50 M", 250, 500, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_50-100 M", 500, 1000, mesAnoCorte, tempFile);
            colocacionOF(okCartera, masterFile, azureFile, fechaCorte, "Colocación_> 100 M", 1000, 10000, mesAnoCorte, tempFile);
            */


            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            deleteTempFile(tempFile);
        } catch (HeadlessException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } /*catch (ParseException e) {
            throw new RuntimeException(e);
        }*/
    }


    public static void nuevosOficinas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        try {


            List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

            List<String> headers = null;
            List<Map<String, String>> datosFiltrados = null;
            List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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


            }
            System.out.println("-----------CREACION TEMPORAL-----------");

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
                runtime();
                waitSeconds(2);

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
                            } else {
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
        } catch (IOException e) {
            throw new RuntimeException("Error interno del proceso", e);
        }
    }

    public static void nuevosOficinasMay30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void nuevosOficinasBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void renovadoOficinas(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void renovadoOficinasMay30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void renovadoOficinasBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, String calificacion, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void oficinasMontoColoc(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
            int valorInicio = valorInic * 100000; // Reemplaza con el valor de inicio del rango
            int valorFin = valorFinal * 100000; // Reemplaza con el valor de fin del rango

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
                        } else {
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

    public static void oficinasMontoColocMay30(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void oficinasMontoColocBE(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String calificacion, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void reestOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "capital");
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
                        } else {
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

    public static void clientesOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "Cliente");
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
                        } else {
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

    public static void operacionesOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String tempFile) throws IOException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "codigo_sucursal");
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
                        } else {
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

    public static void colocacionOF(String okCarteraFile, String masterFile, String azureFile, String fechaCorte, String hoja, int valorInic, int valorFinal, String mesAnoCorte, String tempFile) throws IOException, ParseException {

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "valor_desem");
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
                        } else {
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
