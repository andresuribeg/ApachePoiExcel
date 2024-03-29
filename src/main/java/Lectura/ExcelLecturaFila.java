package Lectura;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

public class ExcelLecturaFila {
    public static void main(String[] args) {
        File archivo = new File("Files/Datos.xlsx");
        try {
            InputStream input = new FileInputStream(archivo);
            XSSFWorkbook libro = new XSSFWorkbook(input);
            XSSFSheet hoja = libro.getSheetAt(1);
            /* XSSFSheet hoja = libro.getSheet("Fila");  Se puede tambien indicar el nombre de la hoja*/
            Row fila = hoja.getRow(1);
            Iterator<Cell> columnas = fila.cellIterator();
            while (columnas.hasNext()) {
                Cell celda = columnas.next();
                if (celda.getCellType() == CellType.STRING) {
                    String valor = celda.getStringCellValue();
                    System.out.println(valor);
                }
                if (celda.getCellType() == CellType.NUMERIC) {
                    double valor = celda.getNumericCellValue();
                    System.out.println(valor);
                }
                if (celda.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(celda)) {
                    SimpleDateFormat formato =new SimpleDateFormat("dd/mm/yyyy");
                    Date fecha = celda.getDateCellValue();
                    System.out.println(fecha);
                }
            }
            input.close();
            libro.close();

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
