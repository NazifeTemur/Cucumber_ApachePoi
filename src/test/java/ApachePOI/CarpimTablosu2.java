package ApachePOI;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;

public class CarpimTablosu2 {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sayfa1");

        int satir = 1;
        for (int i = 1; i < 11; i++) {
            Row yeniSatir = sheet.createRow(satir++);
            int s1 = 1, s2 = 2, s3 = 3, s4 = 4, s5 = 5;
            for (int j = 1; j < 11; j++) {
                Cell yeniHucre = yeniSatir.createCell(s1);
                yeniHucre.setCellValue(j);
                Cell yeniHucre2 = yeniSatir.createCell(s2);
                yeniHucre2.setCellValue("x");
                Cell yeniHucre3 = yeniSatir.createCell(s3);
                yeniHucre3.setCellValue(i);
                Cell yeniHucre4 = yeniSatir.createCell(s4);
                yeniHucre4.setCellValue("=");
                Cell yeniHucre5 = yeniSatir.createCell(s5);
                yeniHucre5.setCellValue(j * i);
                s1 = s1 + 6;  s2 = s2 + 6;  s3 = s3 + 6; s4 = s4 + 6;  s5 = s5 + 6;
            }
        }
                String yeniExcelPath = "src/test/java/ApachePOI/resource/YeniExcel.xlsx";
                FileOutputStream outputStream = new FileOutputStream(yeniExcelPath);
                workbook.write(outputStream);
                workbook.close();
                outputStream.close();
                System.out.println("işlem tamamlandı");
            }
        }
