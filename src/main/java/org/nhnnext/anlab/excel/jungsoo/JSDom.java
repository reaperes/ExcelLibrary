package org.nhnnext.anlab.excel.jungsoo;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class JSDom {
    private String sourceFile = "src/main/resources/source3.xlsx";
    private String destinationFile = "src/main/resources/destination.xlsx";

    private void parse() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(sourceFile)));
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFSheet sheet2 = workbook.getSheetAt(1);

        int maxRow = sheet.getPhysicalNumberOfRows();
        for (int i=0; i<maxRow; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFRow row2 = sheet2.getRow(i);

            if (row2 == null)
                row2 = sheet2.createRow(i);

//            for (int j=0; j<4; j++) {
//                XSSFCell cell = row.getCell(j);
//
//                if (cell != null) {
//                    XSSFCell cell2 = row2.getCell(j);
//                    if (cell2 == null)
//                        cell2 = row2.createCell(j, XSSFCell.CELL_TYPE_STRING);
//
//                    copyCell(cell, cell2);
//                }
//            }

            // special column
            {
                XSSFCell cell = row.getCell(0);
                if (cell != null) {
                    String data = cell.getStringCellValue();
                    String[] parsedString = parsingString(data);

                    // create
                    XSSFCell cell2 = row2.getCell(0);
                    if (cell2 == null)
                        cell2 = row2.createCell(0, XSSFCell.CELL_TYPE_STRING);

                    XSSFCell cell3 = row2.getCell(1);
                    if (cell3 == null)
                        cell3 = row2.createCell(1, XSSFCell.CELL_TYPE_STRING);

                    // copy style
                    cell2.setCellStyle(cell.getCellStyle());
                    cell3.setCellStyle(cell.getCellStyle());

                    // copy value
                    cell2.setCellValue(parsedString[1]);
                    cell3.setCellValue(parsedString[0]);
                    filterStyling(cell3);
                }
            }

            if (i%10 == 0) System.out.println(i + " completed");
        }

        for (int i=0; i<maxRow; i++) {
            XSSFRow row = sheet2.getRow(i);
            XSSFCell cell = row.getCell(1);

            if (cell != null) {
                String data = cell.getStringCellValue();

                Pattern pattern = Pattern.compile("\\s\\([a-zA-Z]");
                Matcher matcher = pattern.matcher(data);

                if (matcher.find()) {
                    if (matcher.find()) {
                        XSSFRichTextString richTextString = cell.getRichStringCellValue();
                        XSSFFont font = new XSSFFont(cell.getCellStyle().getFont().getCTFont());
                        font.setItalic(true);
                        richTextString.applyFont(0, matcher.start(), font);
                        font.setItalic(false);
                    }
                }
            }
        }

        // auto resize
        for (int i=0; i<1; i++)
            sheet2.autoSizeColumn(i);

        workbook.write(new FileOutputStream(new File(destinationFile)));
    }

    private void filterStyling(XSSFCell cell) {
        String text = cell.getStringCellValue();

        Pattern pattern = Pattern.compile("((\\s|\\s\\()[A-Z]|\\s\\(v\\.)");
        Matcher matcher = pattern.matcher(text);

        if (matcher.find()) {
            XSSFRichTextString richTextString = cell.getRichStringCellValue();
            XSSFFont font = new XSSFFont(cell.getCellStyle().getFont().getCTFont());
            font.setItalic(true);
            richTextString.applyFont(0, matcher.start(), font);
            font.setItalic(false);
        }
        else {
            XSSFRichTextString richTextString = cell.getRichStringCellValue();
            XSSFFont font = new XSSFFont(cell.getCellStyle().getFont().getCTFont());
            font.setItalic(true);
            richTextString.applyFont(0, richTextString.length(), font);
            font.setItalic(false);
        }
    }

    private String[] parsingString(String data) {
        String[] ret = new String[2];

        if (checkKukMyeung(data)) {
            ret[0] = data.replaceAll("\\s\\(국명미정\\)", "");
            ret[1] = "(국명미정)";
            return ret;
        }

        String[] strings = data.split(" ");
        StringBuilder sb = new StringBuilder();
        for (String str : strings) {
            if (str.length() < 1)
                continue;

            if (str.charAt(0) < '가' || str.charAt(0) > '힣') {
                if (sb.toString().length() < 1)
                    sb.append(str);
                else
                    sb.append(" " + str);
            }
            else // if korean
                ret[1] = str;
        }
        ret[0] = sb.toString();
        return ret;
    }

    private boolean checkKukMyeung(String data) {
        Pattern pattern = Pattern.compile("\\(국명미정\\)");
        Matcher matcher = pattern.matcher(data);

        return matcher.find();
    }

    private void copyCell(XSSFCell src, XSSFCell dst) throws IllegalArgumentException {
        if (src != null && dst != null) {
            dst.setCellStyle(src.getCellStyle());

            switch (src.getCellType()) {
                case XSSFCell.CELL_TYPE_STRING:
                    dst.setCellValue(src.getStringCellValue());
                    break;

                case XSSFCell.CELL_TYPE_NUMERIC:
                    dst.setCellValue(src.getNumericCellValue());
                    break;

                case XSSFCell.CELL_TYPE_BLANK:
                    break;

                default:
                    throw new IllegalArgumentException("hello, world");
            }
        }
    }

    private void deleteFile() {
        File file = new File(destinationFile);
        if (file != null)
            file.delete();
    }

    public static void main(String[] args) throws Exception {
        JSDom js = new JSDom();
        js.deleteFile();
        js.parse();
    }
}
