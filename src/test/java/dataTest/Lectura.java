package dataTest;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;
import org.utils.MethotsAzureMasterFiles;

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
            FileInputStream fis = new FileInputStream(file1);
            Workbook workbook = new XSSFWorkbook(fis);
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook workbook2 = new XSSFWorkbook(fis2);
            Sheet sheet1 = workbook.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(3);
            List<String> ws1 = MethotsAzureMasterFiles.getWorkSheet(file1, 0);
            List<String> ws2 = dataTest.MethotsAzureMasterFiles.getWorkSheet(file2, 0);

            for (String sheet : ws1){
                for (String sheetName : ws2){

                }
            }
            String s1 = sheet1.getSheetName();
            String s2 = sheet2.getSheetName();
            String sheetName = s2.replaceAll("\\s", "");
            if (!s1.equals(sheetName)){
                for (int indice = 0; indice < workbook.getNumberOfSheets(); indice++) {


                }
                sheet2 = workbook2.getSheetAt(0);
            }else {
                sheet2 = workbook2.getSheetAt(0);
            }



            workbook.close();
            workbook2.close();
            fis.close();
            fis2.close();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
