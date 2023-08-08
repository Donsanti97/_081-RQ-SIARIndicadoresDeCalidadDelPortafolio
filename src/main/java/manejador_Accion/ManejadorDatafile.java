package manejador_Accion;

import com.google.common.base.Splitter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ManejadorDataFile {

    private ManejadorDataFile(){}

    static ManejadorDataFile instanciaManejadorDataFile = null;

    public static void getLoger(String msg){
        Logger loger = Logger.getLogger("Loger");
        try {
            loger.log(Level.WARNING, msg);
        }catch (NullPointerException e){
            loger.log(Level.WARNING, "Error " + e);
        }

    }

    public static ManejadorDataFile getInstanciaManejadorDataFile(){

        if(instanciaManejadorDataFile==null){
            instanciaManejadorDataFile = new ManejadorDataFile();
        }
        return instanciaManejadorDataFile;

    }


    /** Método. Retorna un mapa con la información de un archivo xlsx
     * @param
     * @return List<String>
     * @author Mairon Martinez
     * @since 31/07/2023
     */
    public Map<String, Map<String, String>> datosExcel(String ruta, String nombreHojaExcel) {

        Map <String, Map<String, String>> datosExcel = new HashMap();
        Map <String, String> funciones = new HashMap();

        try{

            File excel = new File(ruta);
            FileInputStream inputStream = new FileInputStream(excel);
            XSSFWorkbook newWorkbook = new XSSFWorkbook(inputStream);
            XSSFSheet hojaExcel = newWorkbook.getSheet(nombreHojaExcel);
            int rowCount = hojaExcel.getLastRowNum() - hojaExcel.getFirstRowNum();
            List<String> titulos = titulosExcel(hojaExcel);


            for (int i = 1; i <= rowCount; i++) {
                XSSFRow row = hojaExcel.getRow(i);


                for (int j = 0; j < row.getLastCellNum(); j++) {
                    funciones.put(titulos.get(j), row.getCell(j).getStringCellValue());
                }
                datosExcel.put(funciones.get("Funcion"), convertWithGuava(funciones.toString().replace(", ", ",").replace("{", "").replace("}", "")));
                funciones.clear();
            }


            return datosExcel;
        }
        catch (IOException e){
            getLoger("Error metodo datosExcel" + e);
            return datosExcel;
        }
    }


    /** Método. Retorna un arreglo con la información de un archivo xlsx
     * @param
     * @return List<String>
     * @author Mairon Martinez
     * @since 31/07/023
     */
    public List<Map<String, String>> consultaDatosExcel(String ruta, String nombreHojaExcel) {

        ArrayList<Map<String, String>> result = new ArrayList();
        Map<String, String> consultaDatosExcel = new HashMap();
        Map <String, String> funciones = new HashMap();

        try{

            File excel = new File(ruta);
            FileInputStream inputStream = new FileInputStream(excel);
            XSSFWorkbook newWorkbook = new XSSFWorkbook(inputStream);
            XSSFSheet hojaExcel = newWorkbook.getSheet(nombreHojaExcel);
            int rowCount = hojaExcel.getLastRowNum() - hojaExcel.getFirstRowNum();
            List<String> titulos = titulosExcel(hojaExcel);

            for (int i = 1; i <= rowCount; i++) {
                XSSFRow row = hojaExcel.getRow(i);

                for (int j = 0; j < row.getLastCellNum(); j++) {
                    funciones.put(titulos.get(j), row.getCell(j).getStringCellValue());
                }
                consultaDatosExcel = convertWithGuava(funciones.toString().replace(", ", ",").replace("{", "").replace("}", ""));
                result.add(consultaDatosExcel);
                funciones.clear();
            }

            return result;
        }
        catch (IOException e){
            getLoger("Error en método consultaDatosExcel" + e);
            return result;
        }
    }

    /** Método. Retorna una lista con los ebcabezados de un excel
     * @param
     * @return List<String>
     * @author Mairon Martinez
     * @since 31/07/023
     */
    public static List<String> titulosExcel(XSSFSheet hojaExcel){
        List<String> titulos = new ArrayList<>();
        for(int i = 0; i<1; i++){
            XSSFRow row = hojaExcel.getRow(i);
            for(int j = 0; j< row.getLastCellNum(); j++) {
                titulos.add(row.getCell(j).getStringCellValue());
            }
        }
        return titulos;
    }
    public static Map<String, String> convertWithGuava(String mapAsString) {
        return Splitter.on(',').withKeyValueSeparator('=').split(mapAsString);
    }
}

