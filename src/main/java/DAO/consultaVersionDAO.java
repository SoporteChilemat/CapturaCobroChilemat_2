package DAO;

import Connect.DbConnection;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class consultaVersionDAO {

    public static ArrayList<String> ConsultaVersion() throws IOException, SQLException {
        DbConnection conex = new DbConnection("");
        ArrayList<String> arrVersion = new ArrayList<>();
        try ( PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT version from dbo.actualizacion");  ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                String version = res.getString("version");
                arrVersion.add(version);
            }
            conex.desconectar();
            return arrVersion;
        } catch (SQLException e) {
            return arrVersion;
        }
    }

    public static byte[] ConsultaPrograma(String version, String nombre) throws IOException, SQLException {
        DbConnection conex = new DbConnection("");
        byte[] bytes = null;
        try ( Statement estatuto = conex.getConnection().createStatement()) {
            ResultSet executeQuery = estatuto.executeQuery("SELECT programa from dbo.actualizacion WHERE version = '" + version + "'");
            if (executeQuery.next()) {
                bytes = executeQuery.getBytes("programa");

                File fileExel = new File(System.getProperty("user.dir") + "\\" + nombre);
                if (!fileExel.exists()) {
                    fileExel.createNewFile();
                }
                FileOutputStream fos = new FileOutputStream(fileExel);
                fos.write(bytes);
                fos.close();
            }
            conex.desconectar();
            return bytes;
        }
    }
}
