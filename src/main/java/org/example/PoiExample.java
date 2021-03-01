package org.example;

import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.helpers.ColumnHelper;

import java.io.IOException;
import java.io.InputStream;

/**
 * poi范例
 * @author zhangwn
 */
public class PoiExample {
    public static final String TEMPLATE_PAH = "template/poi-study.xlsx";

    public static InputStream getInputStream(){
        return Thread.currentThread().getContextClassLoader().getResourceAsStream(TEMPLATE_PAH);
    }

    public static void handle() throws IOException {
        InputStream is = getInputStream();
        XSSFWorkbook book = new XSSFWorkbook(is);
        XSSFSheet sheet = book.getSheetAt(0);

        XSSFRow row = sheet.getRow(0);

        XSSFCell cell = row.getCell(0);
    }

    public static void copy(XSSFSheet dest, XSSFCell cell){
        cell.set
    }

    public static void main(String[] args) throws IOException {
        handle();
    }
}
