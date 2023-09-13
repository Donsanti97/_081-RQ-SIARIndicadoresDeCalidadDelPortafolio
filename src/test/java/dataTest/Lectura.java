package dataTest;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
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

            List<String> encabezados = obtenerEncabezados(sheet);
            List<String> encabezadosSegundoArchivo = encontrarEncabezadosSegundoArchivo(sheet, workbook2);

            for (String encabezado : encabezados) {
                System.out.println("Encabezado: " + encabezado);
            }

            System.out.println("\nEncabezados en el segundo archivo (tomados de la misma columna):");
            for (String encabezado : encabezadosSegundoArchivo) {
                System.out.println("Encabezado: " + encabezado);
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
