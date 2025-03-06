package com.scraping.webscraping;

import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.table.DefaultTableModel;
import org.jsoup.nodes.Document;
import org.jsoup.Jsoup;
import org.jsoup.select.Elements;

/**
 *
 * @author Driscoll
 */
public class CScrap {

    public void ScrapSitioWeb(JTextField urlSitio, JTable resultado) {
        String url = urlSitio.getText();

        if (url.startsWith("https://www.")) {
            url = url;

        } else if (url.startsWith("www.")) {
            url = "https://" + url;
        } else {
            url = "https://www." + url;
        }

        try {
            Document doc = Jsoup.connect(url).get();
            System.out.println("Conexion Exitosa");

            Elements productos = doc.select("a.poly-component__title");
            Elements precio = doc.select("div.poly-component__price");
            Elements precioOferta = doc.select("div.poly-price__current");
            Elements porcentajeOferta = doc.select("span.andes-money-amount__discount");

            DefaultTableModel modelo = new DefaultTableModel(new Object[]{"Titulo", "Precio", "Precio descuento", "Porcentaje de descuento", "Enlace"}, 0);

            resultado.setModel(modelo);

            modelo.setRowCount(0);

            for (int i = 0; i < productos.size(); i++) {

                String titulo = productos.get(i).text();

                String precioNormal = precio.get(i).select("span.andes-money-amount__fraction").first().text();
                String precioNormalString = "$" + precioNormal;

                double precioNormalCast = Double.parseDouble(precioNormal);

                String precioOfertaFinal = precioOferta.get(i).select("span.andes-money-amount__fraction").text();
                String precioOfertaString = "$" + precioOfertaFinal;

                double precioOfertaFinalCast = Double.parseDouble(precioOfertaFinal);

                // No se necesitan
                //String porcentajeDescuento = porcentajeOferta.get(i).select("span.andes-money-amount__discount").text();
                //double porcentajeDescuentoCast = Double.parseDouble(precioOfertaFinal);     
                double porcentajeDescuentoFinal = (precioOfertaFinalCast - precioNormalCast) / (precioNormalCast) * 100;
                porcentajeDescuentoFinal = Math.abs(porcentajeDescuentoFinal);
                int valorDescuento = (int) porcentajeDescuentoFinal;
                String descuentoPorcentaje = valorDescuento + "%";

                String link = productos.get(i).attr("href");

                modelo.addRow(new Object[]{titulo, precioNormalString, precioOfertaString, descuentoPorcentaje, link});

            }

        } catch (Exception e) {
            DefaultTableModel model = new DefaultTableModel(new Object[]{"Error", "Mensaje"}, 0);

            resultado.setModel(model);
            model.setRowCount(0);
            model.addRow(new Object[]{"Error", e.getMessage()});
        }
    }

}
