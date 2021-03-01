package org.example;

import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * poi范例
 * @author zhangwn
 */
public class PoiExample {
    public static final String TEMPLATE_PAH = "template/poi-study.xlsx";
    public static final String OUTPUT_PAH = "E:\\poi-complete.xlsx";

    public static InputStream getInputStream(){
        return Thread.currentThread().getContextClassLoader().getResourceAsStream(TEMPLATE_PAH);
    }

    public static void handle() throws IOException {
        InputStream is = getInputStream();
        XSSFWorkbook book = new XSSFWorkbook(is);
        XSSFSheet sheet = book.getSheetAt(0);

        CellCopyPolicy ccp = new CellCopyPolicy();
        // 复制总和
        XSSFRow row1 = sheet.getRow(1);
        row1.createCell(6);
        CellRangeAddress region = new CellRangeAddress(1, 2, 6, 7);
        sheet.addMergedRegion(region);
        XSSFCell srcCell00 = sheet.getRow(0).getCell(0);
        XSSFCell cell16 = sheet.getRow(1).getCell(6);
        cell16.copyCellFrom(srcCell00, ccp);
        cell16.setCellValue("总和");

        // 两列数据举例
        // copy cells
        for (int i = 0; i < 40; i++) {
            XSSFCell srcCell = sheet.getRow(i).getCell(3);
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.createCell(4);
            // 复制单元格
            cell.copyCellFrom(srcCell, ccp);
            // TODO 设置值
            cell.setCellValue(i);
        }
        output(book);
    }

    public static void output(XSSFWorkbook workbook) throws IOException {
        FileOutputStream out = new FileOutputStream(OUTPUT_PAH);
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws IOException {
        handle();
    }
}
