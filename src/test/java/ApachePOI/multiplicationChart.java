package ApachePOI;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class multiplicationChart {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");

        XSSFCellStyle style10=workbook.createCellStyle();
        style10.setFillForegroundColor(IndexedColors.LAVENDER.getIndex());
        style10.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFCellStyle style1=workbook.createCellStyle();
        style1.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFCellStyle style2=workbook.createCellStyle();
        style2.setFillForegroundColor(IndexedColors.RED.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFCellStyle style3=workbook.createCellStyle();
        style3.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFCellStyle style4=workbook.createCellStyle();
        style4.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        style4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFCellStyle style5=workbook.createCellStyle();
        style5.setFillForegroundColor(IndexedColors.PINK.getIndex());
        style5.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Row baslik0 = sheet.createRow(1);
        Cell yenicell0 = baslik0.createCell(1);
        yenicell0.setCellValue("MULTIPLICATION CHART");
        sheet.addMergedRegion(CellRangeAddress.valueOf("B2:AD2"));
        yenicell0.setCellStyle(style10);
        yenicell0.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

        Row baslik = sheet.createRow(3);
        Cell yenicell = baslik.createCell(3);
        yenicell.setCellValue("ONE");
        Cell yenicell2 = baslik.createCell(9);
        yenicell2.setCellValue("TWO");
        Cell yenicell3 = baslik.createCell(15);
        yenicell3.setCellValue("THREE");
        Cell yenicell4 = baslik.createCell(21);
        yenicell4.setCellValue("FOUR");
        Cell yenicell5 = baslik.createCell(27);
        yenicell5.setCellValue("FIVE");

        Row baslik2 = sheet.createRow(17);
        Cell yenicell6 = baslik2.createCell(3);
        yenicell6.setCellValue("SIX");
        Cell yenicell7 = baslik2.createCell(9);
        yenicell7.setCellValue("SEVEN");
        Cell yenicell8 = baslik2.createCell(15);
        yenicell8.setCellValue("EIGHT");
        Cell yenicell9 = baslik2.createCell(21);
        yenicell9.setCellValue("NINE");
        Cell yenicell10 = baslik2.createCell(27);
        yenicell10.setCellValue("TEN");

        int satir = 4;
        for (int i = 1; i < 11; i++) {
            Row yeniSatir = sheet.createRow(satir++);
            int s1 = 1, s2 = 2, s3 = 3, s4 = 4, s5 = 5;
            for (int j = 1; j < 6; j++) {
                Cell yeniHucre = yeniSatir.createCell(s1);
                yeniHucre.setCellValue(j);
                yeniHucre.setCellStyle(style1);
                yeniHucre.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre2 = yeniSatir.createCell(s2);
                yeniHucre2.setCellValue("x");
                yeniHucre2.setCellStyle(style2);
                yeniHucre2.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre3 = yeniSatir.createCell(s3);
                yeniHucre3.setCellValue(i);
                yeniHucre3.setCellStyle(style3);
                yeniHucre3.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre4 = yeniSatir.createCell(s4);
                yeniHucre4.setCellValue("=");
                yeniHucre4.setCellStyle(style4);
                yeniHucre4.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre5 = yeniSatir.createCell(s5);
                yeniHucre5.setCellValue(j * i);
                yeniHucre5.setCellStyle(style5);
                yeniHucre5.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                s1 = s1 + 6;  s2 = s2 + 6;  s3 = s3 + 6;  s4 = s4 + 6;  s5 = s5 + 6;
            }
        }
        int satir2 = 18;
        for (int i = 1; i < 11; i++) {
            Row yeniSatir = sheet.createRow(satir2++);
            int s1 = 1, s2 = 2, s3 = 3, s4 = 4, s5 = 5;
            for (int j = 6; j < 11; j++) {
                Cell yeniHucre = yeniSatir.createCell(s1);
                yeniHucre.setCellValue(j);
                yeniHucre.setCellStyle(style1);
                yeniHucre.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre2 = yeniSatir.createCell(s2);
                yeniHucre2.setCellValue("x");
                yeniHucre2.setCellStyle(style2);
                yeniHucre2.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre3 = yeniSatir.createCell(s3);
                yeniHucre3.setCellValue(i);
                yeniHucre3.setCellStyle(style3);
                yeniHucre3.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre4 = yeniSatir.createCell(s4);
                yeniHucre4.setCellValue("=");
                yeniHucre4.setCellStyle(style4);
                yeniHucre4.getCellStyle().setAlignment(HorizontalAlignment.CENTER);

                Cell yeniHucre5 = yeniSatir.createCell(s5);
                yeniHucre5.setCellValue(j * i);
                yeniHucre5.setCellStyle(style5);
                yeniHucre5.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                s1 = s1 + 6; s2 = s2 + 6;  s3 = s3 + 6; s4 = s4 + 6;  s5 = s5 + 6;
            }
        }

              String yeniExcelPath = "src/test/java/ApachePOI/resource/newExcel.xlsx";
                FileOutputStream outputStream = new FileOutputStream(yeniExcelPath);
                workbook.write(outputStream);
                workbook.close();
                outputStream.close();
                System.out.println("DONE");
            }
        }
