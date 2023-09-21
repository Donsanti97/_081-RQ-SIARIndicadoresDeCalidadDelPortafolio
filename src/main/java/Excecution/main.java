package Excecution;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import static utils.MethotsAzureMasterFiles.*;


public class main {
    public static void main(String[] args) {
        try {
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Azure a analizar");
            String file1 = getDocument();
            JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro a analizar");
            String file2 = getDocument();

            if (file1 != null && file2 != null) {
                System.out.println("Ruta del archivo Excel seleccionado: " + file1);
                System.out.println("Ruta del archivo Excel seleccionado: " + file2);
            } else {
                System.out.println("No se seleccionó ningún archivo.");
            }

            FileInputStream fis = new FileInputStream(file1);
            Workbook workbook = new XSSFWorkbook(fis);
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook workbook2 = new XSSFWorkbook(fis2);
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(3);

            List<String> nameSheets1 = getWorkSheet(file1, 0);
            List<String> nameSheets2 = getWorkSheet(file2, 3);

            List<String> encabezados1 = null;
            List<String> encabezados2 = null;

            List<Map<String, String>> valoresEncabezados1 = null;
            List<Map<String, String>> valoresEncabezados2 = null;


            for (String sheets : nameSheets1) {
                System.out.print("SheetName: ");
                System.out.println(sheets);
                encabezados1 = getHeaders(file1, "CER150");
                //System.out.println("Headers: ");
                for (String headers : encabezados1) {
                    //System.out.print(headers + "||");
                    valoresEncabezados1 = getValuebyHeader(file1, sheets);
                }
            }
            System.out.println("------------------------------------------------------------------------------------------");

            valoresEncabezados1 = obtenerValoresPorFilas(sheet1, encabezados1);
            for (Map<String, String> map : valoresEncabezados1) {
                System.out.println("Fila: ");
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    System.out.println("Headers1: " + entry.getKey() + ", value: " + entry.getValue());
                }
            }

            System.out.println("------------------------------------------------------");
            for (String sheets2 : nameSheets2) {
                System.out.println("SheetName2: " + sheets2);
                encabezados2 = getHeadersMasterfile(sheet1, sheet2);
                for (String headers : encabezados2) {
                    System.out.println("Headers2: " + headers);
                }
            }

            System.out.println("-------------------------------------------------------------------------------------");
            valoresEncabezados2 = obtenerValoresPorFilas(sheet2, encabezados2);
            for (Map<String, String> map : valoresEncabezados2) {
                System.out.println("Fila2: ");
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    System.out.println("Headers2: " + entry.getKey() + ", Value: " + entry.getValue());
                }
            }

            System.out.println("---------------------------------------------------------------------------------------");
            for (String e1 : encabezados1) {
                for (String e2 : encabezados2) {
                    if (e1.equals(e2)) {
                        System.out.println("equals" + e1 + ", " + e2);
                    } else {
                        System.out.println("No equals");
                    }
                }
            }


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
   /* @Test
    public static void test() {
        try {
            FileInputStream fis = new FileInputStream(file1);
            Workbook workbook = new XSSFWorkbook(fis);
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook workbook2 = new XSSFWorkbook(fis2);
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(3);

            List<String> nameSheets1 = getWorkSheet(file1, 0);
            List<String> nameSheets2 = getWorkSheet(file2, 3);

            List<String> encabezados1 = null;
            List<String> encabezados2 = null;

            List<Map<String, String>> valoresEncabezados1 = null;
            List<Map<String, String>> valoresEncabezados2 = null;


            for (String sheets : nameSheets1) {
                System.out.print("SheetName: ");
                System.out.println(sheets);
                encabezados1 = getHeaders(file1, "CER150");
                //System.out.println("Headers: ");
                for (String headers : encabezados1) {
                    //System.out.print(headers + "||");
                    valoresEncabezados1 = getValuebyHeader(file1, sheets);
                }
            }
            System.out.println("------------------------------------------------------------------------------------------");

            valoresEncabezados1 = obtenerValoresPorFilas(sheet1, encabezados1);
            for (Map<String, String> map : valoresEncabezados1) {
                System.out.println("Fila: ");
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    System.out.println("Headers1: " + entry.getKey() + ", value: " + entry.getValue());
                }
            }

            System.out.println("------------------------------------------------------");
            for (String sheets2 : nameSheets2) {
                System.out.println("SheetName2: " + sheets2);
                encabezados2 = getHeadersMasterfile(sheet1, sheet2);
                for (String headers : encabezados2) {
                    System.out.println("Headers2: " + headers);
                }
            }

            System.out.println("-------------------------------------------------------------------------------------");
            valoresEncabezados2 = obtenerValoresPorFilas(sheet2, encabezados2);
            for (Map<String, String> map : valoresEncabezados2) {
                System.out.println("Fila2: ");
                for (Map.Entry<String, String> entry : map.entrySet()) {
                    System.out.println("Headers2: " + entry.getKey() + ", Value: " + entry.getValue());
                }
            }

            System.out.println("---------------------------------------------------------------------------------------");
            for (String e1 : encabezados1) {
                for (String e2 : encabezados2) {
                    if (e1.equals(e2)) {
                        System.out.println("equals" + e1 + ", " + e2);
                    } else {
                        System.out.println("No equals");
                    }
                }
            }


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }*/

}
