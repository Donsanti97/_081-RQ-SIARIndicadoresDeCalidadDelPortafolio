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

import javax.swing.*;
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import static dataTest.FunctionsApachePoi.*;
import static dataTest.FunctionsApachePoi.crearNuevaHojaExcel;
import static org.utils.MethotsAzureMasterFiles.getDocument;


public class NumericValues {



    @Test
    public static void TEST() {
        String excelFilePathTest = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleOfTheMiddleTestData.xlsx";
        String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel

        IOUtils.setByteArrayMaxOverride(300000000);
        System.out.println("URL " + excelFilePathTest);
        Scanner scanner = new Scanner(System.in);
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure");
        String azureFile = getDocument();
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        String masterFile = getDocument();
        JOptionPane.showMessageDialog(null, "Seleccione el archivo OkCartera");
        String okCartera = getDocument();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola el número del mes y año de corte del archivo OkCartera sin espacios (Ejemplo: 02/2023 (febrero/2023))");
        String mesAñoCorte = mostrarCuadroDeTexto();
        JOptionPane.showMessageDialog(null, "ingrese a continuación en la consola la fecha de corte del archivo OkCartera sin espacios (Ejemplo: 30/02/2023)");
        String fechaCorte = mostrarCuadroDeTexto();



        try {
            findFields(okCartera, masterFile, azureFile, mesAñoCorte, fechaCorte);

        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }


    public static void findFields(String okCarteraFile, String masterFile, String azureFile, String mesAñoCorte, String fechaCorte) throws IOException, ParseException {

        //String excelFilePath = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TablaDinamica.xlsx"; // Reemplaza con la ruta de tu archivo Excel
        //String okCarteraFile = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\MiddleTestData.xlsx";

        IOUtils.setByteArrayMaxOverride(300000000);

        List<String> sheetNames = obtenerNombresDeHojas(okCarteraFile);

        List<String> headers = null;
        List<Map<String, String>> datosFiltrados = null;
        List<String> camposDeseados = Arrays.asList("codigo_sucursal", "plazo");
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
            String fechaInicio = "01/" + mesAñoCorte;
            String fechaFin = "31/" + mesAñoCorte;
            Date rangoInicio = new SimpleDateFormat("dd/MM/yyyy").parse(fechaInicio);
            Date rangoFin = new SimpleDateFormat("dd/MM/yyyy").parse(fechaFin);


            // Filtrar los datos por el campo y el rango especificados
            datosFiltrados = obtenerValoresDeEncabezados(okCarteraFile, sheetName, campoFiltrar, valorInicio, valorFin, "fecha_inicio_cre", rangoInicio, rangoFin);

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

            for (Map.Entry<String, String> entryOkCartera : resultado.entrySet()){
                System.out.println("Valor único del primer encabezado: " + entryOkCartera.getKey());
                System.out.println("Suma correspondiente: " + entryOkCartera.getValue());
                System.out.println();

                List<Map<String, String>> resultados = analisisMasterFile(azureFile, masterFile, fechaCorte);
                for (Map<String, String> datos : resultados) {
                    for (Map.Entry<String, String> entryMasterFile : datos.entrySet()) {
                        System.out.println("KEY: " + entryMasterFile.getKey() + ", VALUE: " + entryMasterFile.getValue());
                    }
                }

            }

            System.out.println("---------------------- CREACION TABLA DINAMICA comercialPzoPerc08");


            tablasDinamicasApachePoi(nuevaHojaFilePath, camposDeseados.get(0), camposDeseados.get(1), "SUMA");

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
