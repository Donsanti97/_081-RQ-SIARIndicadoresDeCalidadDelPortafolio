package dataTest;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static dataTest.FunctionsApachePoi.*;


public class Lectura {



    @Test
    public static void analisisArchivos(){
        try {
            String file1 = System.getProperty("user.dir") + "\\documents\\initialDocument\\Historico Cartera Comercial.xlsx";
            String file2 = System.getProperty("user.dir") + "\\documents\\finalDocument\\Historico Cartera COMERCIAL por OF.xlsx";
            System.out.println(file1);
            FileInputStream excelFile = new FileInputStream(file1);
            FileInputStream excelFile2 = new FileInputStream(file2);


            Workbook workbook2 = new XSSFWorkbook(excelFile2);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0); // Puedes especificar el índice de la hoja que desees procesar
            Sheet sheet2 = workbook2.getSheetAt(3);

            List<String> encabezados = obtenerEncabezados(sheet);
            List<String> encabezadoTest = buscarValorEnColumna(sheet2, 0, encabezados.get(0));
            Map<String, List<String>> valoresPorEncabezado1 = obtenerValoresPorEncabezado(sheet, encabezados);
            //assert encabezadoTest != null;
            Map<String, List<String>> valoresPorEncabezado2 = obtenerValoresPorEncabezado(sheet2, encabezadoTest);

            for (String encabezado : encabezados) {
                System.out.println("Encabezado: " + encabezado);
            }
            for (String encabezado : encabezadoTest){
                System.out.println("EncabezadoTest: " + encabezado);
            }

            for (String encabezado : encabezados) {
                List<String> valores = valoresPorEncabezado1.get(encabezado);
                System.out.println("Encabezado: " + encabezado);
                for (String valor : valores) {
                    System.out.println("  Valor: " + valor);
                }
            }

            /*List<String> sheetNames1 = obtenerNombresDeHojas(file1);
            List<String> headers1 = null;

            List<String> sheetNames2 = obtenerNombresDeHojas(file2, 3);
            List<String> headers2 = null;

            for (String sheetName1 : sheetNames1) {
                headers1 = obtenerEncabezados(file1, sheetName1);
                for (String sheetName2 : sheetNames2) {
                    System.out.println("SheetNames1: " + sheetName1 + "\nSheetNames2: " + sheetName2);

                    headers2 = obtenerEncabezados(file2, sheetName2);
                    System.out.println("Headers1: " + headers1 + "\n Headers2: " + headers2);
                }
            }*/

            workbook.close();
            workbook2.close();
            excelFile.close();
            excelFile2.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
