
package Manejo;

import java.awt.Desktop;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.CopyOption;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import javax.swing.JOptionPane;


public class Manejo {
    
    public static void renameFichero(String nombre1, String nombre2) {
        if (nombre2.length() < 2) {
            File fichero = new File(nombre1);
            File fichero2 = new File(nombre1);
        } else {
            File fichero = new File(nombre1);
            File fichero2 = new File(nombre2);
            boolean success = fichero.renameTo(fichero2);
            if (!success) {
                System.out.println("Error intentando cambiar el nombre de fichero");
                //fichero.delete();
            }
        }
    }

    public static String directorioActual() {
        File miDir = new File(".");
        try {
            return miDir.getCanonicalPath();
        } catch (Exception e) {
            return e.toString();
        }
    }

    public static void fileMove(String sourceFile, String destinationFile, String nombre) throws IOException {

        Path FROM = Paths.get(sourceFile);
        Path TO = Paths.get(destinationFile);
        CopyOption[] options = new CopyOption[]{
            //StandardCopyOption.REPLACE_EXISTING,
            StandardCopyOption.COPY_ATTRIBUTES
        };
        Files.copy(FROM, TO, options);
        //File file = new File(directorioActual() + "\\" + nombre + ".jpg");
        System.out.println(directorioActual() + "\\" + nombre + ".jpg");
        //file.delete();
    }

    public static void crearCarpeta(String variable) {
        try {
            File directorio = new File(directorioActual() + "\\" + variable);
            if (!directorio.exists()) {

                directorio.mkdirs();
            }
        } catch (Exception e) {
            System.err.println(e);
        }
    }

    public static int ciclo() {
        int num = 0;
        return num;
    }

    public static boolean abrirFichero(String ruta) {
        boolean bool = false;
        String directorioActual = directorioActual();
        Desktop ficheroAEjecutar = Desktop.getDesktop();
        try {
            File file = new File(directorioActual + "\\" + ruta + ".jpg");
            if (file.length() != 0) {
                ficheroAEjecutar.open(new File(directorioActual + "\\" + ruta + ".jpg"));
                bool = true;
            } else {
                file.delete();
            }
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
        return bool;
    }  
}
