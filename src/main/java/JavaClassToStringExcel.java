import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JavaClassToStringExcel {
    public static void main(String[] args) {
        String javaClassFile = "C:\\Users\\ABC\\git\\APMIS-Project\\apmis-flow\\src\\main\\java\\com\\cinoteck\\application\\views\\MainLayout.java";
        String excelFile = "C:\\Users\\ABC\\Projects\\copyStrings\\src\\main\\resources\\excel\\mainman.xlsx";

        List<String> strings = extractStringsFromJavaClass(javaClassFile);

        try {
            writeStringsToExcel(strings, excelFile);
            System.out.println("Strings copied to Excel successfully.");
        } catch (IOException e) {
            System.err.println("Error writing to Excel: " + e.getMessage());
        }
    }

    private static List<String> extractStringsFromJavaClass(String javaClassFile) {
        List<String> strings = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(javaClassFile)) {
            // Assuming that the file contains valid Java code
            int content;
            StringBuilder stringBuilder = new StringBuilder();
            boolean inString = false;

            while ((content = fis.read()) != -1) {
                char character = (char) content;

                if (character == '"' && !inString) {
                    inString = true;
                } else if (character == '"' && inString) {
                    inString = false;
                    strings.add(stringBuilder.toString());
                    stringBuilder.setLength(0);
                } else if (inString) {
                    stringBuilder.append(character);
                }
            }
        } catch (IOException e) {
            System.err.println("Error reading Java class file: " + e.getMessage());
        }

        return strings;
    }

    private static void writeStringsToExcel(List<String> strings, String excelFile) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Strings");

        int rowIdx = 0;
        for (String string : strings) {
            Row row = sheet.createRow(rowIdx++);
            Cell cell = row.createCell(0);
            cell.setCellValue(string);
        }

        try (FileOutputStream fos = new FileOutputStream(excelFile)) {
            workbook.write(fos);
        }

        workbook.close();
    }
}
