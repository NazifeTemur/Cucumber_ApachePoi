package ApachePOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CarpimTablosu {
    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sayfa1");

            int satirno=0;
            for (int i = 1; i < 11; i++) {
                for (int j = 1; j < 11; j++) {
                Row yeniSatir=sheet.createRow(satirno++);
                Cell yeniHucre = yeniSatir.createCell(0);
                yeniHucre.setCellValue(i);
                Cell yeniHucre2 = yeniSatir.createCell(1);
                yeniHucre2.setCellValue("x");
                Cell yeniHucre3 = yeniSatir.createCell(2);
                yeniHucre3.setCellValue(j);
                Cell yeniHucre4 = yeniSatir.createCell(3);
                yeniHucre4.setCellValue("=");
                Cell yeniHucre5 = yeniSatir.createCell(4);
                yeniHucre5.setCellValue(j*i);
            }
                Row yeniSatir2=sheet.createRow(satirno++);
            }

            String yeniExcelPath = "src/test/java/ApachePOI/resource/YeniExcel.xlsx";
            FileOutputStream outputStream = new FileOutputStream(yeniExcelPath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            System.out.println("işlem tamamlandı");
        }
    }
