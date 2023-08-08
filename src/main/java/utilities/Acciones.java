package utilities;

import org.openqa.selenium.WebDriver;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

public class Acciones {
    public Acciones(WebDriver webDriver) {
        super();
    }

    //Este método mueve un archivo de una carpeta a otra
    public static void moveFile(File analizedFile, File toDirectory) throws IOException {
        File file = new File(System.getProperty("user.dir")+"\\finalDocument\\"+ analizedFile.getName());
        System.out.println(file.getAbsoluteFile()+"--------");
        File directory = new File(createDirectory(toDirectory.getName()) + "\\" + file.getName());
        System.out.println(directory.getAbsoluteFile()+"-------");
        if (directory.exists()) {
            if (file.exists()) {
                Files.move(file.toPath(), directory.toPath());
            }
        }
    }

    public static File createDirectory(String dirName){
        File newDir = new File(System.getProperty("user.dir")+"\\documents\\" + dirName);
        if(newDir.mkdir()){
            System.out.println("Directory has been created");
        }else {
            System.out.println("Directory cannot be created");
        }

        return newDir;
    }



}
