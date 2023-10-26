package dataTest;

import dataTest.historicoCarteraComercialPorOF.HistoricoCarteraComercialPorOF;
import dataTest.historicoCarteraConsumoPorOF.HistoricoCarteraConsumoPorOF;
import dataTest.historicoCarteraMicrocreditoPorOF.HistoricoCarteraMicrocreditoPorOF;
import org.testng.annotations.Test;

import javax.swing.*;

import java.io.File;

import static org.utils.MethotsAzureMasterFiles.getDocument;

public class Start {
    @Test
    public void excecution(){
        JOptionPane.showMessageDialog(null, "Seleccione el archivo Maestro");
        String masterFile = getDocument();

        try {
            assert masterFile != null;
            File file = new File(masterFile);
            System.out.println(file.getName());
            String fileName = file.getName().toLowerCase();
            System.out.println(fileName);
            if (fileName.contains("comercial")){
                HistoricoCarteraComercialPorOF.configuracion(masterFile);
            } else if (fileName.contains("consumo")) {
                HistoricoCarteraConsumoPorOF.configuracion(masterFile);
            } else if (fileName.contains("microcredito")) {
                HistoricoCarteraMicrocreditoPorOF.configuracion(masterFile);
            }else {
                System.out.println("EL ARCHIVO SELECCIONADO NO TIENE ANÁLISIS ASIGNADO");
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

    }
}
