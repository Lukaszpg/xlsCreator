package commons;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class XlsFileCreator<T> {

    private Class<T> clazz;

    public XlsFileCreator(Class<T> clazz) {
        this.clazz = clazz;
    }

    public void createFile(List<T> series, String path, String fileName)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException, IOException {

        //tworzę plik excel
        HSSFWorkbook workbook = new HSSFWorkbook();

        // tworzę arkusz w pliku
        HSSFSheet sheet = workbook.createSheet(fileName);

        //ustawiam czcionki
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        //zapisuję styl czcionki do arkusza
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        //kolekja nazw kolumn w arkuszu
        List<String> columnsTitles = new ArrayList<>();

        //iteracja po klasie przekazanej do pola 'clazz'. Odczytuję wszystkie zadeklarowane pola.
        for (Field f : clazz.getDeclaredFields()) {
            columnsTitles.add(f.getName());
        }

        //zapisuję do struktury pliku excel odczytane powyżej pola klasy jako nagłówki kolumn.
        Row headerRow = sheet.createRow(0);
        for (int i = 0; i < columnsTitles.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columnsTitles.get(i));
            cell.setCellStyle(headerCellStyle);
        }

        // test odczytu metod zaczynających się na 'get'. Tylko test.
        columnsTitles.forEach(t -> System.out.println("get" + t.substring(0, 1).toUpperCase() + t.substring(1)));

        //zapis danych i wywoływanie metod 'get'.
        for (int i = 1; i < series.size(); i++) {

            HSSFRow row = sheet.createRow(i);

            for (int j = 0; j < columnsTitles.size(); j++) {

                HSSFCell cell = row.createCell(j);
                Method method = series.get(i)
                        .getClass()
                        .getMethod("get" + columnsTitles.get(j)
                                .substring(0, 1)
                                .toUpperCase() + columnsTitles.get(j)
                                .substring(1));

                Object result = method.invoke(series.get(i));
                cell.setCellValue(String.valueOf(result));
            }
        }

        //ustawianie auto szerokości kolumn.
        for (int i = 0; i < columnsTitles.size(); i++) {
            sheet.autoSizeColumn(i);
        }

        //dodawanie unikalnej nazwy do nazwy pliku. Przykład: persons_12948573628.xls
        long mills = System.currentTimeMillis();
        String file = path + fileName + "_" + mills + ".xls";

        //zapis pliku
        workbook.write(new File(file));
        workbook.close();
    }

}
