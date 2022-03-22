package DAO;

import Clases.DocumentoCobranza;
import java.io.IOException;
import java.sql.SQLException;
import Clases.Totales;
import Connect.DbConnection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.ArrayList;
import javax.swing.JOptionPane;

public class TotalesDAO {

    public static ArrayList<Totales> consultaTotales(String bd) throws IOException, SQLException {
        ArrayList<Totales> arrTotales = new ArrayList<>();
        DbConnection conex = new DbConnection(bd);
        try ( PreparedStatement consulta = conex.getConnection().prepareStatement("SELECT * FROM ingresos.totales");  ResultSet res = consulta.executeQuery()) {
            while (res.next()) {
                Totales totales = new Totales();
                totales.setFechas(res.getString("fechas"));
                totales.setSelected(res.getInt("selected"));
                totales.setComentario(res.getString("comentario"));
                arrTotales.add(totales);
            }
            conex.desconectar();
        }
        return arrTotales;
    }

    public static void registraTotales(Totales totales, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        try {
            Statement estatuto = conex.getConnection().createStatement();
            estatuto.executeUpdate("INSERT INTO ingresos.totales (fechas, selected, comentario) VALUES ('"
                    + totales.getFechas() + "', '"
                    + totales.getSelected() + "', '"
                    + totales.getComentario() + "')"
            );
            estatuto.close();
            conex.desconectar();
        } catch (Exception ex) {

        }
    }

    public static void actualizaSelected(Totales totales, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        String fechas = totales.getFechas();
        int selected = totales.getSelected();
        System.out.println("fechas " + fechas + " selected " + selected);

        try ( Statement estatuto = conex.getConnection().createStatement()) {
            estatuto.executeUpdate("UPDATE ingresos.totales SET selected = '"
                    + totales.getSelected() + "' WHERE fechas = '"
                    + totales.getFechas().trim() + "'");
            JOptionPane.showMessageDialog(null, "Se ha actualizado Exitosamente", "Información", JOptionPane.INFORMATION_MESSAGE);
            conex.desconectar();
        }
    }

    public static void actualizaComentario(Totales totales, String bd) throws IOException, SQLException {
        DbConnection conex = new DbConnection(bd);
        String fechas = totales.getFechas();
        String comentario = totales.getComentario();
        System.out.println("fechas " + fechas + " comentario " + comentario);

        try ( Statement estatuto = conex.getConnection().createStatement()) {
            estatuto.executeUpdate("UPDATE ingresos.totales SET comentario = '"
                    + totales.getComentario().trim() + "' WHERE fechas = '"
                    + totales.getFechas().trim() + "'");
            JOptionPane.showMessageDialog(null, "Se ha actualizado Exitosamente", "Información", JOptionPane.INFORMATION_MESSAGE);
            conex.desconectar();
        }
    }
}
