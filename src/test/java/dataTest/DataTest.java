package dataTest;

import com.aspose.cells.*;
import org.testng.annotations.Test;

import java.io.File;
public class DataTest {


    @Test(description = "Consulta de campos tipo String")
    public static void consultaDatos(){
        try {


            //File file = new File("4. Historico Cartera COMERCIAL por OF.xlsx");
            //File newDir = new File("procesedDocuments");

            String file1 =System.getProperty("user.dir") + "\\documents\\initialDocument\\4. Historico Cartera COMERCIAL por OF.xlsx";
            String file2 =System.getProperty("user.dir") + "\\documents\\finalDocument\\4. Historico Cartera COMERCIAL por OF.xlsx";

            String test1 = "4. Historico Cartera COMERCIAL por OF.xlsx";
            String test2 = "4. Historico Cartera COMERCIAL por OF.xlsx";

            System.out.println(file1+"------");
            System.out.println(file2+"-------");



            if (getInformation(test1).equals(getInformation(test2))){
                System.out.println("Compatibles");
            }else {
                System.out.println("Nel pastel");
            }



            System.out.println("------------------------------------------------------------------------------");
            Workbook wb = new Workbook(file1);

            WorksheetCollection collection = wb.getWorksheets();

            for (int worksheetIndex = 0 ; worksheetIndex < collection.getCount(); worksheetIndex++){
                Worksheet worksheet = collection.get(worksheetIndex);

                System.out.println("Worksheet: " + worksheet.getName());

                int rows = worksheet.getCells().getMaxDataRow();
                int cols = worksheet.getCells().getMaxDataColumn();

                for (int i = 0; i < rows; i++) {

                    for (int j = 0; j < cols; j++) {
                        if (worksheet.getCells().get(i, j).getValue() != null) {
                            System.out.println(worksheet.getCells().get(176, 3).toString()/*getDisplayStringValue().getValue()*/ + "||");
                            break;
                        }
                    }
                    System.out.println(" ");
                }
            }






        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * Al implementar un filtro de carga personalizado, permite abilitar un filtro para hojas Excel individuales
     * @Description: Se crea una clase y un método dentro de la clase con el propósito de filtrar datos numéricos
     * String y celdas con formulas
     * **/
    class CustomLoadFilter extends LoadFilter {
        public void starSheet(Worksheet worksheet) {

            if (worksheet.getName().equals("Nombre de la hoja")) {
                //Filtra todos los valores, y a la vez los numéricos
                this.setLoadDataFilterOptions(LoadDataFilterOptions.ALL & LoadDataFilterOptions.CELL_NUMERIC);
            }

            if(worksheet.getName().equals("Nombre de la hoja")){
                //Filtra todos los valores, y a la vez los String
                this.setLoadDataFilterOptions(LoadDataFilterOptions.ALL & LoadDataFilterOptions.CELL_STRING);
            }

            if(worksheet.getName().equals("Nombre de la hoja")){
                //Filtra todos los valores y a la vez formatos condicionados
                this.setLoadDataFilterOptions(LoadDataFilterOptions.ALL & LoadDataFilterOptions.CONDITIONAL_FORMATTING);
            }
        }
    }

    public void Run() throws Exception{
        String dataDir = System.getProperty("user.dir") + "\\documents\\initialDocument\\4. Historico Cartera COMERCIAL por OF.xlsx";

        //Filtrar hojaExcel usando filtro de carga personalizado
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadFilter(new CustomLoadFilter());

        //Cargar el archivo Excel con el filtro anterior
        Workbook wb = new Workbook(dataDir, loadOptions);

        //Toma la imagen de todas las hojas una por una
        for (int i = 0; i < wb.getWorksheets().getCount(); i++) {
            //Accede a las hojas en el indice i
            Worksheet ws = wb.getWorksheets().get(i);

            //Crea imagen o imprime opciones, queremos la imagen de la hoja entera
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.setOnePagePerSheet(true);
            options.setImageType(ImageType.PNG);

            //Convierte la hoja en imagen
            SheetRender sr = new SheetRender(ws, options);
        }
    }

    public static String getInformation(String fileName) throws Exception {

        String string = "Valiste berenjena";

        String file = System.getProperty("user.dir") + "\\documents\\initialDocument\\" + fileName;

        Workbook wb = new Workbook(file);

        WorksheetCollection collection = wb.getWorksheets();

        for (int worksheetIndex = 0 ; worksheetIndex < collection.getCount(); worksheetIndex++){
            Worksheet worksheet = collection.get(worksheetIndex);

            System.out.println("Worksheet: " + worksheet.getName());

            int rows = worksheet.getCells().getMaxDataRow();
            int cols = worksheet.getCells().getMaxDataColumn();

            for (int i = 0; i < rows; i++) {

                for (int j = 0; j < cols; j++) {
                    if (worksheet.getCells().get(i, j).getValue() != null && worksheet.getCells().get(i, j).isNumericValue()) {
                        System.out.println(worksheet.getCells().get(176, 3).getDisplayStringValue() + "||");
                        string = worksheet.getCells().get(176, 3).getDisplayStringValue() + "||";
                        return string;
                        /*if (worksheet.getCells().get(176, 8).isNumericValue()){
                            System.out.println("Es un numero pape");
                            return string;
                        }else if (!worksheet.getCells().get(176, 8).isNumericValue()){
                            System.out.println(worksheet.getCells().get(176, 8).getDisplayStringValue() + "||");
                            string = worksheet.getCells().get(176, 8).getDisplayStringValue() + "||";
                            return string;
                        }*/
                    }
                }
                System.out.println(" ");
            }
        }


        return string;
    }


}
