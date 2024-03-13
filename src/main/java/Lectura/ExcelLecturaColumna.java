package Lectura;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Iterator;

public class ExcelLectura {
    public static void main(String[] args) {
        File archivo = new File("Datos.xlsx");
        try {
            InputStream input = new FileInputStream(archivo);
            XSSFWorkbook libro = new XSSFWorkbook(input);
            XSSFSheet hoja = libro.getSheetAt(0);
//            Row fila = hoja.getRow(1) asi traería la fila específica
            Iterator<Row> filas = hoja.rowIterator();

            Cell columna =null;
            while (filas.hasNext()) {
                columna =filas.next().getCell(0);

                System.out.println(columna.getStringCellValue());

            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}