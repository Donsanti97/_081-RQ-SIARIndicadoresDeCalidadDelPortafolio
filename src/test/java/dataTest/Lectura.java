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
import static dataTest.FunctionsApachePoi.*;


public class Lectura {



    @Test
    public static void analisisArchivos(){
        try {
            FileInputStream excelFile = new FileInputStream(System.getProperty("user.dir") + "\\documents\\initialDocument\\Historico Cartera Comercial.xlsx");
            FileInputStream excelFile2 = new FileInputStream(System.getProperty("user.dir") + "\\documents\\finalDocument\\Historico Cartera COMERCIAL por OF.xlsx");
            Workbook workbook2 = new XSSFWorkbook(excelFile2);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet sheet = workbook.getSheetAt(0); // Puedes especificar el índice de la hoja que desees procesar
            Sheet sheet2 = workbook2.getSheetAt(3);

            List<String> encabezados = obtenerEncabezados(sheet);
            List<String> encabezadoTest = buscarValorEnColumna(sheet2, 0, encabezados.get(0));

            for (String encabezado : encabezados) {
                System.out.println("Encabezado: " + encabezado);
            }
            for (String encabezado : encabezadoTest){
                System.out.println("EncabezadoTest: " + encabezado);
            }

            List<String> sheetNames1 = obtenerNombresDeHojas(excelFile.toString());
            List<String> headers1 = null;

            List<String> sheetNames2 = obtenerNombresDeHojas(excelFile2.toString(), 3);
            List<String> headers2 = null;

            for (String sheetName1 : sheetNames1) {
                for (String sheetName2 : sheetNames2) {
                    System.out.println("SheetNames1: " + sheetName1 + "\nSheetNames2: " + sheetName2);
                    headers1 = obtenerEncabezados(excelFile.toString(), sheetName1);
                    headers2 = obtenerEncabezados(excelFile2.toString(), sheetName2);




                }
            }

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
