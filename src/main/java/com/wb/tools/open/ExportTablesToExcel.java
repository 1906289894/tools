package com.wb.tools.open;

import cn.hutool.core.util.StrUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.sql.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class ExportTablesToExcel {
    public static List<String> filterTables = new ArrayList<>(Arrays.asList(
            "m_user", "t_assess_car", "t_assess_scheme",
            "t_attachment", "t_attachment_temp", "t_business_module",
            "t_car", "t_car_brand", "t_car_category", "t_car_model",
            "t_car_series", "t_commercial_car_tire_financing", "t_customer_info",
            "t_dealer_business_sub_type", "t_dealer_business_type",
            "t_dealer_car_model", "t_dealer_car_to_common_record",
            "t_dealer_login_log", "t_dealer_project", "t_dealer_project_dealer",
            "t_financing_loan_order_batch_credit", "t_financing_loan_order_subtask",
            "t_financing_scheme", "t_financing_scheme_fuel_order",
            "t_financing_scheme_insurance_order", "t_financing_scheme_order_car",
            "t_financing_scheme_spare_parts_order", "t_financing_scheme_tire_order",
            "t_pv_partner_vc_customer_info", "t_pv_partner_vc_customer_task",
            "t_pv_partner_vc_info", "t_pv_partner_vc_loanorder",
            "t_pv_partner_vc_order_subtask", "t_remote_verify_order",
            "t_vc_apply", "t_vc_assign_account", "t_vc_customer", "t_vc_user"
    ));

    public static void main(String[] args) {
        String jdbcUrl = "jdbc:postgresql://192.168.1.40:5432/dev?currentSchema=tj_psbc_carloan";
        String username = "pg";
        String password = "123.abc";

        try (Connection connection = DriverManager.getConnection(jdbcUrl, username, password)) {
            DatabaseMetaData metaData = connection.getMetaData();
            ResultSet tables = metaData.getTables(null, "tj_psbc_carloan", "%", new String[]{"TABLE"});

            Workbook workbook = new XSSFWorkbook();

            while (tables.next()) {
                String tableName = tables.getString("TABLE_NAME");
                String tableComment = getTableComment(connection, tableName);
                //过滤不需要的表，匹配一个集合
                if (filterTable(tableName, tableComment)) {
                    continue;
                }


                Sheet sheet = workbook.createSheet(getSheetName(tableName, tableComment));

                ResultSet columns = metaData.getColumns(null, "tj_psbc_carloan", tableName, "%");

                int rowNum = 0;
                Row headerRow = sheet.createRow(rowNum++);
                String[] headers = {"字段序号", "表英文名", "表中文名", "字段英文名", "字段中文名", "字段类型", "长度", "精度", "是否主键", "外键", "是否可以为空", "缺省值", "取值范围", "业务取值说明", "当前状态", "备注", "企业级数据字典数据项编号", "企业级数据字典匹配结果"};


                // 创建样式
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                headerCellStyle.setBorderBottom(BorderStyle.THIN);
                headerCellStyle.setBorderTop(BorderStyle.THIN);
                headerCellStyle.setBorderRight(BorderStyle.THIN);
                headerCellStyle.setBorderLeft(BorderStyle.THIN);

                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]);
                    cell.setCellStyle(headerCellStyle);
                }

                // 创建样式
                CellStyle dataCellStyle = workbook.createCellStyle();
                dataCellStyle.setBorderBottom(BorderStyle.THIN);
                dataCellStyle.setBorderTop(BorderStyle.THIN);
                dataCellStyle.setBorderRight(BorderStyle.THIN);
                dataCellStyle.setBorderLeft(BorderStyle.THIN);

                int index = 1; // 初始化字段序号
                while (columns.next()) {
                    String columnName = columns.getString("COLUMN_NAME");
                    String columnComment = getColumnComment(connection, tableName, columnName);
                    int dataType= columns.getInt("DATA_TYPE");
                    int columnSize = columns.getInt("COLUMN_SIZE");
                    // 是否为空
                    boolean nullable = columns.getInt("NULLABLE") == DatabaseMetaData.columnNullable;

                    Row row = sheet.createRow(rowNum++);
                    Cell cell0 = row.createCell(0);
                    cell0.setCellValue(index++);
                    cell0.setCellStyle(dataCellStyle); // 字段序号

                    Cell cell1 = row.createCell(1);
                    cell1.setCellValue(tableName);
                    cell1.setCellStyle(dataCellStyle); // 表英文名

                    Cell cell2 = row.createCell(2);
                    cell2.setCellValue(StrUtil.isBlank(tableComment) ? " " : tableComment);
                    cell2.setCellStyle(dataCellStyle); // 表中文名

                    Cell cell3 = row.createCell(3);
                    cell3.setCellValue(columnName);
                    cell3.setCellStyle(dataCellStyle); // 字段英文名

                    Cell cell4 = row.createCell(4);
                    cell4.setCellValue(columnComment);
                    cell4.setCellStyle(dataCellStyle); // 字段中文名

                    Cell cell5 = row.createCell(5);
                    cell5.setCellValue(mapDataType(dataType));
                    cell5.setCellStyle(dataCellStyle); // 字段类型

                    Cell cell6 = row.createCell(6);
                    cell6.setCellValue(columnSize);
                    cell6.setCellStyle(dataCellStyle); // 长度
                    // 精度

                    Cell cell7 = row.createCell(7);
                    cell7.setCellStyle(dataCellStyle);

                    Cell cell8 = row.createCell(8);
                    cell8.setCellStyle(dataCellStyle);
                    if ("id".equalsIgnoreCase(columnName)){
                        cell8.setCellValue("Y");
                    }else{
                        cell8.setCellValue("N");
                    }

                    Cell cell9 = row.createCell(9);
                    cell9.setCellStyle(dataCellStyle);

                    Cell cell10 = row.createCell(10);
                    cell10.setCellStyle(dataCellStyle);
                    if (nullable) {
                        cell10.setCellValue("Y");
                    } else {
                        cell10.setCellValue("N");
                    }

                    // 其他字段信息填充

                    for (int j = 11; j < headers.length; j++) {
                        Cell cell = row.createCell(j);
                        cell.setCellStyle(dataCellStyle); // 其他字段样式
                    }
                }
            }

            // 将 workbook 写入到文件
            FileOutputStream fileOut = new FileOutputStream("tables_info1.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("Excel 文件已创建");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static boolean filterTable(String tableName, String tableComment) {
        if (!filterTables.contains(tableName)) {
            System.out.println("过滤 tableName: " + tableName);
            return true;
        }
        System.out.println("add tableName: " + tableName + "_" + tableComment);
        return false;
    }

    private static String mapDataType(int dataType) {
        switch (dataType) {
            case Types.BIT:
                return "bit";
            case Types.BOOLEAN:
                return "boolean";
            case Types.TINYINT:
                return "tinyint";
            case Types.SMALLINT:
                return "smallint";
            case Types.INTEGER:
                return "integer";
            case Types.BIGINT:
                return "bigint";
            case Types.FLOAT:
                return "float";
            case Types.REAL:
                return "real";
            case Types.DOUBLE:
                return "double";
            case Types.NUMERIC:
                return "numeric";
            case Types.DECIMAL:
                return "decimal";
            case Types.CHAR:
                return "char";
            case Types.VARCHAR:
                return "varchar";
            case Types.LONGVARCHAR:
                return "longvarchar";
            case Types.DATE:
                return "date";
            case Types.TIME:
                return "time";
            case Types.TIMESTAMP:
                return "timestamp";
            // 其他数据类型的映射
            default:
                return "other";
        }
    }


    // 获取表的中文注释
    private static String getTableComment(Connection connection, String tableName) throws SQLException {
        String query = "SELECT obj_description((SELECT oid FROM pg_class WHERE relname = ? LIMIT 1), 'pg_class')";

        try (PreparedStatement statement = connection.prepareStatement(query)) {
            statement.setString(1, tableName);

            try (ResultSet resultSet = statement.executeQuery()) {
                if (resultSet.next()) {
                    return resultSet.getString(1);
                } else {
                    return " "; // 如果未找到表的中文注释，则返回空字符串
                }
            }
        }
    }


    // 获取表的字段的中文注释
    private static String getColumnComment(Connection connection, String tableName, String columnName) throws SQLException {
        String query = "SELECT description " +
                "FROM pg_description " +
                "JOIN pg_attribute ON pg_attribute.attnum = pg_description.objsubid " +
                "JOIN pg_class ON pg_class.oid = pg_attribute.attrelid " +
                "WHERE pg_class.relname = ? AND pg_attribute.attname = ? LIMIT 1";

        try (PreparedStatement statement = connection.prepareStatement(query)) {
            statement.setString(1, tableName);
            statement.setString(2, columnName);

            try (ResultSet resultSet = statement.executeQuery()) {
                if (resultSet.next()) {
                    return resultSet.getString(1);
                } else {
                    return " "; // 如果未找到字段的中文注释，则返回空字符串
                }
            }
        }
    }

    // 创建 sheet 名称
    private static String getSheetName(String tableName, String tableComment) {
        if (StrUtil.isBlank(tableComment) || tableComment.trim().isEmpty()) {
            return tableName;
        } else {
            return tableName + "_" + tableComment;
        }
    }
}