package com.example.dataToExcel;


import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.dao.DataAccessException;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.ResultSetExtractor;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;



@RestController

public class DataToExcelController {
    public static int columnCount;
    @Autowired
    JdbcTemplate jdbc;
    @RequestMapping("/insert")

    public void createFile(){
        String sql = "SELECT * FROM City"; // you can specify any query here
       // String sql = "SELECT Name, CountryCode FROM City";
        //String sql = "SELECT * FROM Country";
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("city");
        //Util util= new Util();
        jdbc.query(sql, new ResultSetExtractor() {
            @Override

            public String extractData(ResultSet rs) throws SQLException, DataAccessException {

                //util.writeHeaderLine(rs, sheet);
                //util.writeDataLines(rs, workbook, sheet);
                ResultSetMetaData metaData = rs.getMetaData();
                int numberOfColumns = metaData.getColumnCount();

                Row headerRow = sheet.createRow(0);

                // exclude the first column as it the ID
                for (int i = 1; i <= numberOfColumns; i++) {
                    String columnName = metaData.getColumnName(i);
                    Cell headerCell = headerRow.createCell(i - 1);
                    headerCell.setCellValue(columnName);
                }
                int rowCount = 1;

                while (rs.next()) {
                    Row row = sheet.createRow(rowCount++);

                    for (int i = 1; i <= numberOfColumns; i++) {
                        Object valueObject = rs.getObject(i);

                        Cell cell = row.createCell(i - 1);

                        if (valueObject instanceof Boolean)
                            cell.setCellValue((Boolean) valueObject);
                        else if (valueObject instanceof Double)
                            cell.setCellValue((double) valueObject);
                        else if (valueObject instanceof Float)
                            cell.setCellValue((float) valueObject);
                        else if (valueObject instanceof Integer)
                            cell.setCellValue((int) valueObject);
                        /*else if (valueObject instanceof Date) {
                            cell.setCellValue((Date) valueObject);
                            util.formatDateCell(workbook, cell);}*/
                         else cell.setCellValue((String) valueObject);

                    }

                }
                try {

                        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH-mm-ss");
                        String dateTimeInfo = dateFormat.format(new Date());
                        String ExcelPath ="output".concat(String.format("_%s.xlsx", dateTimeInfo));

                        FileOutputStream outputStream = new FileOutputStream(ExcelPath);
                    workbook.write(outputStream);
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }

                return "File created ";
            }

        });

    }
}
