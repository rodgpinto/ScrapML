package com.scraping.webscraping;

import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jsoup.nodes.Document;
import org.jsoup.Jsoup;
import org.jsoup.select.Elements;
import java.io.*;

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
            System.out.println("Conexión Exitosa");

            Elements productos = doc.select("a.poly-component__title");
            Elements precio = doc.select("div.poly-component__price");
            Elements precioOferta = doc.select("div.poly-price__current");

            DefaultTableModel modelo = new DefaultTableModel(new Object[]{"Título", "Precio", "Precio descuento", "Porcentaje de descuento", "Enlace"}, 0);
            resultado.setModel(modelo);

            for (int i = 0; i < productos.size(); i++) {
                String titulo = productos.get(i).text();

                String precioNormal = precio.get(i).select("span.andes-money-amount__fraction").first().text();
                String precioNormalString = "$" + precioNormal;

                double precioNormalCast = Double.parseDouble(precioNormal);

                String precioOfertaFinal = precioOferta.get(i).select("span.andes-money-amount__fraction").text();
                String precioOfertaString = "$" + precioOfertaFinal;

                double precioOfertaFinalCast = Double.parseDouble(precioOfertaFinal);

                double porcentajeDescuentoFinal = (precioOfertaFinalCast - precioNormalCast) / precioNormalCast * 100;
                porcentajeDescuentoFinal = Math.abs(porcentajeDescuentoFinal);
                int valorDescuento = (int) porcentajeDescuentoFinal;
                String descuentoPorcentaje = valorDescuento + "%";

                String link = productos.get(i).attr("href");

                modelo.addRow(new Object[]{titulo, precioNormalString, precioOfertaString, descuentoPorcentaje, link});
            }

            // Obtener ruta del escritorio y actualizar el archivo
            String rutaEscritorio = System.getProperty("user.home") + "/Desktop/";
            String nombreArchivo = rutaEscritorio + "OfertasMercadoLibre.xlsx";
            exportarAExcel(resultado, nombreArchivo);

        } catch (Exception e) {
            DefaultTableModel model = new DefaultTableModel(new Object[]{"Error", "Mensaje"}, 0);
            resultado.setModel(model);
            model.addRow(new Object[]{"Error", e.getMessage()});
        }
    }

    public void exportarAExcel(JTable table, String rutaArchivo) {
        Workbook workbook;
        Sheet sheet;

        File file = new File(rutaArchivo);
        boolean archivoExiste = file.exists();

        try {
            if (archivoExiste) {
                // Si el archivo ya existe, lo abrimos
                FileInputStream fileInputStream = new FileInputStream(file);
                workbook = new XSSFWorkbook(fileInputStream);
                sheet = workbook.getSheetAt(0);  // Usamos la primera hoja
                fileInputStream.close();
            } else {
                // Si no existe, creamos un nuevo archivo
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet("Ofertas");

                // Agregar encabezados si el archivo es nuevo
                Row headerRow = sheet.createRow(0);
                DefaultTableModel model = (DefaultTableModel) table.getModel();
                for (int i = 0; i < model.getColumnCount(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(model.getColumnName(i));
                    cell.setCellStyle(getHeaderCellStyle(workbook));
                }
            }

            // Insertar los nuevos datos al final de la tabla
            DefaultTableModel model = (DefaultTableModel) table.getModel();
            int ultimaFila = sheet.getLastRowNum() + 1;

            for (int i = 0; i < model.getRowCount(); i++) {
                Row row = sheet.createRow(ultimaFila + i);
                for (int j = 0; j < model.getColumnCount(); j++) {
                    row.createCell(j).setCellValue(model.getValueAt(i, j).toString());
                }
            }

            // Autoajustar el tamaño de las columnas
            for (int i = 0; i < model.getColumnCount(); i++) {
                sheet.autoSizeColumn(i);
            }

            // Guardar el archivo actualizado
            FileOutputStream fileOut = new FileOutputStream(file);
            workbook.write(fileOut);
            fileOut.close();

            workbook.close();
            System.out.println("Archivo Excel actualizado correctamente en: " + rutaArchivo);

        } catch (IOException e) {  // Eliminamos InvalidFormatException
            e.printStackTrace();
        }
    }

    private CellStyle getHeaderCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
}
