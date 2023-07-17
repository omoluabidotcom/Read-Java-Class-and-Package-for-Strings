import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class JavaPackageToStringExcel {
    public static void main(String[] args) {
        String packagePath = "package local path";
        String excelFile = "excelsheet local destination";

        List<String> strings = extractStringsFromJavaPackage(packagePath);

        try {
            writeStringsToExcel(strings, excelFile);
            System.out.println("Strings copied to Excel successfully.");
        } catch (IOException e) {
            System.err.println("Error writing to Excel: " + e.getMessage());
        }
    }

    private static List<String> extractStringsFromJavaPackage(String packagePath) {
        List<String> strings = new ArrayList<>();

        File packageDirectory = new File(packagePath);
        if (!packageDirectory.exists() || !packageDirectory.isDirectory()) {
            System.err.println("Invalid package directory path.");
            return strings;
        }

        File[] files = packageDirectory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isFile() && file.getName().endsWith(".java")) {
                    strings.addAll(extractStringsFromJavaClass(file.getAbsolutePath()));
                }
            }
        }

        return strings;
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
