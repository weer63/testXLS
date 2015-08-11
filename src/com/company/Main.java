package com.company;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {
	    // write your code here
        int i=0;
        String xlsFileName = "D:\\00000000000000\\report.xls";
        HSSFWorkbook xlsReportBook = new HSSFWorkbook();
        HSSFSheet xlsSheet = xlsReportBook.createSheet("Report");
        xlsSheet.setColumnWidth(0,10000);
        xlsSheet.setColumnWidth(1,7000);
        xlsSheet.setColumnWidth(2, 5000);
        HSSFRow xlsRowHead = xlsSheet.createRow((short) (i++));
        HSSFCell cellA1 =  xlsRowHead.createCell(0);
        cellA1.setCellValue("Company Name");
        HSSFCell cellA2= xlsRowHead.createCell(1);
        cellA2.setCellValue("Accounts Code");
        HSSFCell cellA3 = xlsRowHead.createCell(2);
        cellA3.setCellValue("Posting Type");

        HSSFCellStyle cellStyle = xlsReportBook.createCellStyle();
        cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        cellStyle.setBorderBottom(HSSFBorderFormatting.BORDER_MEDIUM);
        cellStyle.setBorderTop(HSSFBorderFormatting.BORDER_MEDIUM);
        cellStyle.setBorderLeft(HSSFBorderFormatting.BORDER_MEDIUM);
        cellStyle.setBorderRight(HSSFBorderFormatting.BORDER_MEDIUM);
        HSSFFont hSSFFont = xlsReportBook.createFont();
        hSSFFont.setBold(true);
        hSSFFont.setFontHeight((short)250);
        cellStyle.setFont(hSSFFont);

        cellA1.setCellStyle(cellStyle);
        cellA2.setCellStyle(cellStyle);
        cellA3.setCellStyle(cellStyle);
        FileOutputStream xlsFile = new FileOutputStream(xlsFileName);
        xlsReportBook.write(xlsFile);
        xlsFile.close();
    }
}
