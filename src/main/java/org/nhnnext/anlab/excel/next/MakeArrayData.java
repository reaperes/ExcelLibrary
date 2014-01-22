package org.nhnnext.anlab.excel.next;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class MakeArrayData {
    private final String sourceFile = "src/main/resources/list.xlsx";

    public static void main(String[] args) throws Exception {
        MakeArrayData js = new MakeArrayData();
        js.make();
    }

    private void make() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(sourceFile)));
        XSSFSheet sheet1 = workbook.getSheetAt(0);
        XSSFSheet sheet2 = workbook.getSheetAt(1);

        int i=0;
        Iterator<Row> rows = sheet2.iterator();
        StringBuilder sb = new StringBuilder();
        while (rows.hasNext()) {
            i++;
            String name = rows.next().getCell(0).getStringCellValue();
            name = name.replaceAll(".(?=.)", "$0 ");
            sb.append("\"" + name + "\", ");
        }

        System.out.println(i);
        System.out.println(sb.toString());
    }
}