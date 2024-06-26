package com.example.excelUpload.service;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelTOSQL {public static void generateSQL(String tableName, String excelFilePath, String commandType) throws IOException {
    List<List<String>> data = readFile(excelFilePath);
    String sql = "";

    switch (commandType.toUpperCase()) {
        case "INSERT":
            sql = generateInsertSQL(tableName, data);
            break;
        case "UPDATE":
            sql = generateUpdateSQL(tableName, data);
            break;
        case "DELETE":
            sql = generateDeleteSQL(tableName, data);
            break;
        default:
            System.out.println("Invalid command type. Please use Insert, Update, or Delete.");
            return;
    }

    writeSQLToFile(sql, tableName + ".sql");
}

    private static List<String[]> readExcel(String excelFilePath) {
        List<String[]> data = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                int numCells = row.getPhysicalNumberOfCells();
                String[] rowData = new String[numCells];
                for (int i = 0; i < numCells; i++) {
                    rowData[i] = row.getCell(i).toString();
                }
                data.add(rowData);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }

    private static String generateInsertSQL(String tableName, List<List<String>> data) {
        StringBuilder sql = new StringBuilder();
        List<String> headers = data.get(0);
        for (int i = 1; i < data.size(); i++) {
            sql.append("INSERT INTO ").append(tableName).append(" (");
            for (int j = 0; j < headers.size(); j++) {
                sql.append(headers.get(j));
                if (j < headers.size() - 1) sql.append(", ");
            }
            sql.append(") VALUES (");
            for (int j = 0; j < data.get(i).size(); j++) {
                sql.append("'").append(data.get(i).get(j)).append("'");
                if (j < data.get(i).size() - 1) sql.append(", ");
            }
            sql.append(");\n");
        }
        return sql.toString();
    }

    private static String generateUpdateSQL(String tableName, List<List<String>> data) {
        // Assumes the first column is the unique identifier for the update
        StringBuilder sql = new StringBuilder();
        List<String> headers = data.get(0);
        for (int i = 1; i < data.size(); i++) {
            sql.append("UPDATE ").append(tableName).append(" SET ");
            for (int j = 1; j < headers.size(); j++) {
                sql.append(headers.get(j)).append(" = '").append(data.get(i).get(j)).append("'");
                if (j < headers.size() - 1) sql.append(", ");
            }
            sql.append(" WHERE ").append(headers.get(0)).append(" = '").append(data.get(i).get(0)).append("';\n");
        }
        return sql.toString();
    }

    private static String generateDeleteSQL(String tableName, List<List<String>> data) {
        // Assumes the first column is the unique identifier for the delete
        StringBuilder sql = new StringBuilder();
        List<String> headers = data.get(0);
        for (int i = 1; i < data.size(); i++) {
            sql.append("DELETE FROM ").append(tableName).append(" WHERE ")
                    .append(headers.get(0)).append(" = '").append(data.get(i).get(0)).append("';\n");
        }
        return sql.toString();
    }

    private static void writeSQLToFile(String sql, String filePath) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            writer.write(sql);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static List<List<String>> readFile(String file) throws IOException {
        Workbook workbook = null;
        InputStream inputStream = null;
        List<List<String>> data = new ArrayList<>();
        try {
            File file1= new File(file);
//            inputStream = file1.getInputStream();
            workbook = new XSSFWorkbook(file1);
            workbook.setMissingCellPolicy(Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            Iterator<Row> rows = sheet.iterator();
//            rows.next();
            while (rows.hasNext()) {
                Row currentRow = rows.next();
                List<String> rowData = new ArrayList<>();
                for (int col = 0; col < currentRow.getLastCellNum(); col++) {
                    Cell column = currentRow.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (column != null) {
                        switch (column.getCellType()) {
                            case STRING:
                                rowData.add(column.getStringCellValue());
                                break;
                            case NUMERIC:
                                rowData.add(String.valueOf(column.getNumericCellValue()));
                                break;
                            default:
                                rowData.add("NULL");
                                break;
                        }
                    } else {
                        rowData.add("NULL");
                    }

                }

                data.add(rowData);
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        } finally {
            if (workbook != null) {
                workbook.close();
            }
            if (inputStream != null) {
                inputStream.close();
            }
        }

        return data;
    }
}
