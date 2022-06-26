package com.example.demo.helper;

import com.example.demo.model.Cars;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class excelHelper {
    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static String[] HEADERs = { "Id", "Name", "Count"};
    static String SHEET = "Tutorials";

    public static boolean hasExcelFormat(MultipartFile file) {

        if (!TYPE.equals(file.getContentType())) {
            return false;
        }

        return true;
    }

    public static ByteArrayInputStream carsToExcel(List<Cars> cars){
        Workbook workbook = new XSSFWorkbook();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            Sheet sheet = workbook.createSheet(SHEET);

            // Header
            Row headerRow = sheet.createRow(0);

            for (int col = 0; col < HEADERs.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(HEADERs[col]);
            }

            int rowIdx = 1;
            for (Cars car : cars) {
                Row row = sheet.createRow(rowIdx++);

                row.createCell(0).setCellValue(car.getId());
                row.createCell(1).setCellValue(car.getName());
                row.createCell(2).setCellValue(car.getCount());
            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
        }
    }

    public static ByteArrayInputStream jsonCarsToExcel(String json) {
        try {
            ObjectMapper mapper = new ObjectMapper();
            Workbook workbook = new XSSFWorkbook();
            ObjectNode jsonData = (ObjectNode) mapper.readTree(json);
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            Iterator<String> sheetItr = jsonData.fieldNames();
            while (sheetItr.hasNext()) {
                String sheetName = sheetItr.next();
                Sheet sheet = workbook.createSheet(sheetName);

                ArrayNode sheetData = (ArrayNode) jsonData.get(sheetName);
                ArrayList<String> headers = new ArrayList<String>();

                CellStyle headerStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                headerStyle.setFont(font);

                Row header = sheet.createRow(0);
                Iterator<String> it = sheetData.get(0).fieldNames();
                int headerIdx = 0;
                while (it.hasNext()) {
                    String headerName = it.next();
                    headers.add(headerName);
                    Cell cell=header.createCell(headerIdx++);
                    cell.setCellValue(headerName);
                    cell.setCellStyle(headerStyle);
                }

                for (int i = 0; i < sheetData.size(); i++) {
                    ObjectNode rowData = (ObjectNode) sheetData.get(i);
                    Row row = sheet.createRow(i + 1);
                    for (int j = 0; j < headers.size(); j++) {
                        String value = rowData.get(headers.get(j)).asText();
                        row.createCell(j).setCellValue(value);
                    }
                }

                for (int i = 0; i < headers.size(); i++) {
                    sheet.autoSizeColumn(i);
                }

            }

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());
        } catch (IOException e) {
            throw new RuntimeException("fail to import data to Excel file: " + e.getMessage());
        }
    }

    public static List<Cars> excelToCars(InputStream is) {
        try {
            Workbook workbook = new XSSFWorkbook(is);

            Sheet sheet = workbook.getSheet(SHEET);
            Iterator<Row> rows = sheet.iterator();

            List<Cars> cars = new ArrayList<Cars>();

            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Cars car = new Cars();

                int cellIdx = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    switch (cellIdx) {
                        case 0:
                            car.setId((long) currentCell.getNumericCellValue());
                            break;

                        case 1:
                            car.setName(currentCell.getStringCellValue());
                            break;

                        case 2:
                            car.setCount((long) currentCell.getNumericCellValue());
                            break;

                        default:
                            break;
                    }

                    cellIdx++;
                }

                cars.add(car);
            }


            return cars;
        } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
    }

}
