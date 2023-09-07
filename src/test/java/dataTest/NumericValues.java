package dataTest;

import com.aspose.cells.*;
import org.testng.annotations.Test;

public class NumericValues {



    @Test
    public static void createDinamicTableAposeCells() throws Exception {
        String file1 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TestData.xlsx";

        // Crear una instancia de un objeto Workbook
        Workbook workbook = new Workbook(file1);

// Accede a la ficha
        Worksheet sheet2 = workbook.getWorksheets().get(1);

// Obtenga la colección de tablas dinámicas en la hoja
        PivotTableCollection pivotTables = sheet2.getPivotTables();

// Agregar una tabla dinámica a la hoja de trabajo
        int index = pivotTables.add("=Data!A1:DT180936", "DX4", "PivotTable1");

// Acceda a la instancia de la tabla dinámica recién agregada
        PivotTable pivotTable = pivotTables.get(index);

// Mostrar los totales generales
        pivotTable.setRowGrand(true);
        pivotTable.setColumnGrand(true);

// Establecer que el informe de tabla dinámica se formatea automáticamente
        pivotTable.setAutoFormat(true);

// Establezca el tipo de formato automático de la tabla dinámica.
        pivotTable.setAutoFormatType(PivotTableAutoFormatType.REPORT_6);

// Arrastre el primer campo al área de la fila.
        pivotTable.addFieldToArea(PivotFieldType.ROW, 8);//Codigo_sucursal

// Arrastre el tercer campo al área de la fila.
        pivotTable.addFieldToArea(PivotFieldType.ROW, 2);

// Arrastre el segundo campo al área de la fila.
        pivotTable.addFieldToArea(PivotFieldType.ROW, 1);

// Arrastre el cuarto campo al área de la columna.
        pivotTable.addFieldToArea(PivotFieldType.COLUMN, 3);

// Arrastre el quinto campo al área de datos.
        pivotTable.addFieldToArea(PivotFieldType.DATA, 12);//Modalidad
        pivotTable.addFieldToArea(PivotFieldType.DATA, 15);//dias_de_mora

// Establecer el formato de número del primer campo de datos
        pivotTable.getDataFields().get(0).setNumber(7);

// Guarde el archivo de Excel
        workbook.save("pivotTable.xls");

    }


    /*@Test
    public static void modificarMacro() throws Exception {
        String file1 = System.getProperty("user.dir") + "\\documents\\procesedDocuments\\TestData.xlsm";
        Workbook workbook = new Workbook(file1);

        try {
            Worksheet worksheet = workbook.getWorksheets().get(0);
            VbaModuleCollection modules = workbook.getVbaProject().getModules();
            VbaModule module = modules.get(0);
            System.out.println(module.getName());
            String code = module.getCodes();

            if (code.contains("Holis")){
                code = code.replace("Holis", "Holis2");
                module.setCodes(code);
            }

            /*for (int i = 0; i < modules.getCount(); i++) {
                VbaModule module = modules.get(i);
                String code = module.getCodes();

                if (code.contains("This is test message.")) {
                    code = code.replace("This is test message.", "This is Apose.Cells message.");
                }
            }
            workbook.save("TestData.xlsm");

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }*/
}
