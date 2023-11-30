package dataTest.historicoCarteraBrutaPorOF_LC;

import javax.swing.*;
import java.awt.*;

import static dataTest.FunctionsApachePoi.*;
import static dataTest.MethotsAzureMasterFiles.*;

public class HistoricoCarteraBrutaPorOF_LC {
    //197 hojas

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

            JOptionPane.showMessageDialog(null, "Espere un momento el análisis puede ser demorado...");
            waitMinutes(5);

            waitMinutes(12);

            JOptionPane.showMessageDialog(null, "Archivos analizados correctamente...");
            waitSeconds(10);

            deleteTempFile(tempFile);
        } catch (HeadlessException e) {
            throw new RuntimeException(e);
        }
    }
}
