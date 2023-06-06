package zadatak1;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class Zadatak1 {
    public static void main(String[] args) throws IOException {
        String relativePath = "NumericData.xlsx";

        try {
            readAndWriteData(relativePath);
        } catch (FileNotFoundException exception) {
            System.out.println("Invalid path");
        } catch (NullPointerException nullPointerException) {
            System.out.println("Data does not exist");
        } catch (IOException ioException) {
            System.out.println("Invalid Excel file");
        }
    }

    public static void readAndWriteData(String relativePath) throws FileNotFoundException, IOException {

        FileInputStream fileInputStream = new FileInputStream(relativePath);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheet("Numbers");

        XSSFSheet sheet2 = workbook.createSheet("Average");

        XSSFRow row = sheet.getRow(0);
        XSSFCell cell = row.getCell(0);

        int rowCount = sheet.getPhysicalNumberOfRows();
        int cellCount = row.getPhysicalNumberOfCells();

        for (int i = 0; i < rowCount; i++) {

            row = sheet.getRow(i);
            XSSFRow row2 = sheet2.createRow(i);
            double sum = 0;
            double avg = 0;

            for (int j = 0; j < cellCount; j++) {
                cell = row.getCell(j);
                double number = cell.getNumericCellValue();
                System.out.print(number + " | ");
                sum += number;
                avg = sum / 5;
                XSSFCell cell2 = row2.createCell(0);
                cell2.setCellValue(avg);
            }
            System.out.println();
            System.out.println(avg);
        }

        FileOutputStream fileOutputStream = new FileOutputStream(relativePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();

    }
}
