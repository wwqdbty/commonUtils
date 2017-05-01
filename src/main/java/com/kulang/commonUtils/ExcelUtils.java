package com.kulang.commonUtils;

import com.google.common.base.Preconditions;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Created by 1 on 2017/2/12.
 */
public class ExcelUtils {
    public static String getReplaceChar(String sourceStr) {
        if (StringUtils.isBlank(sourceStr)) {
            return "";
        }

        if (sourceStr.contains("（")) {
            sourceStr = sourceStr.replace("（", "(");
        }
        if (sourceStr.contains("）")) {
            sourceStr = sourceStr.replace("）", ")");
        }
        if (sourceStr.contains("—")) {
            sourceStr = sourceStr.replace("—", "-");
        }
        if (sourceStr.contains("－")) {
            sourceStr = sourceStr.replace("－", "-");
        }
        if (sourceStr.contains("／")) {
            sourceStr = sourceStr.replace("／", "/");
        }
        if (sourceStr.contains(" ")) {
            sourceStr = sourceStr.replace(" ", "");
        }


        return sourceStr;
    }

    /**
     * 解析Excel, 返回key为sheet号, value为内容的key自然顺序的TreeMap结构
     * @param isExcel07
     * @param inputStream
     * @return
     */
    public static Map<Integer, List<String[]>> parseExcel(boolean isExcel07, InputStream inputStream) {
        Preconditions.checkNotNull(inputStream);

        Workbook workbook = null;
        try {
            if (isExcel07) {
                workbook = new XSSFWorkbook(inputStream);
            } else {
                workbook = new HSSFWorkbook(inputStream);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();

            return null;
        } catch (IOException e) {
            e.printStackTrace();

            return null;
        }

        Map<Integer, List<String[]>> mapListSheetData = new TreeMap<>();
        // sheet数
        int iSheetCount = workbook.getNumberOfSheets();
        for (int iSheetIndex = 0; iSheetIndex < iSheetCount; iSheetIndex++) {
            Sheet sheet = workbook.getSheetAt(iSheetIndex);

            List<String[]> listSheetData = getSheetData(sheet);
            // 无论如何返回数据
            if (listSheetData == null) {
                listSheetData = Collections.EMPTY_LIST;
            }

            mapListSheetData.put(iSheetIndex, listSheetData);
        }

        return mapListSheetData;
    }

    /**
     * 获取Sheet数据
     * @param sheet
     * @return
     */
    private static List<String[]> getSheetData(final Sheet sheet) {
        if (sheet == null) {
            return null;
        }

        // sheet中数据集合
        List<String[]> listSheetData = new ArrayList<>();

        int iRowCount = sheet.getPhysicalNumberOfRows();
        // 仅认为从第二行开始为正式数据
        for (int iRowIndex = 1; iRowIndex < iRowCount; iRowIndex++) {
            Row row = sheet.getRow(iRowIndex);
            String[] arrRowData = getRowData(row);
            // 无数据的空行,　则跳过
            if (arrRowData == null) {
                continue;
            }
            listSheetData.add(arrRowData);
        }

        return listSheetData;
    }

    /**
     * 获取行数据
     * @param row
     * @return
     */
    private static String[] getRowData(final Row row) {
        if (row == null) {
            return null;
        }
        // 列数
        int iCellCount = row.getPhysicalNumberOfCells();
        if (iCellCount == 0) {
            return null;
        }

        // 行的所有列为空标识
        boolean bAllEmptyFlag = true;

        String[] arrRowData = new String[iCellCount];

        for (int iCellIndex = 0; iCellIndex < iCellCount; iCellIndex++) {
            Cell cell = row.getCell(iCellIndex);
            if (cell == null) {
                continue;
            }

            String cellValue = getCellValue(cell);
            arrRowData[iCellIndex] = cellValue.trim();

            // 重置空标志位
            if (StringUtils.isNotBlank(cellValue)) {
                bAllEmptyFlag = false;
            }
        }

        if (bAllEmptyFlag) {
            return null;
        }
        return arrRowData;
    }

    /**
     * 获取Cell值, 所有类型将被转换为String
     * @param cell
     * @return
     */
    private static String getCellValue(final Cell cell) {
        if (cell == null) {
            return null;
        }

        String strCellValue = "";

        int cellType = cell.getCellType();
        switch (cellType) {
            // 文本
            case Cell.CELL_TYPE_STRING:
                strCellValue = cell.getStringCellValue();
                break;

            // 数字或日期
            case Cell.CELL_TYPE_NUMERIC:
                // 日期
                if (DateUtil.isCellDateFormatted(cell)) {
                    strCellValue = new DateTime(cell.getDateCellValue()).toString("yyyy-MM-dd HH:mm:ss");
//                    strCellValue = DateUtils.formatDate(cell.getDateCellValue(), "yyyy-MM-dd HH:mm:ss");
                    break;
                }

                // 数字
                strCellValue = String.valueOf(cell.getNumericCellValue());
                if (strCellValue.contains("E")) {
                    strCellValue = String.valueOf(new Double(cell.getNumericCellValue()).longValue());
                }
                break;

            // 布尔型
            case Cell.CELL_TYPE_BOOLEAN:
                strCellValue = String.valueOf(cell.getBooleanCellValue());
                break;

            // 空白
            case Cell.CELL_TYPE_BLANK:
                strCellValue = cell.getStringCellValue();
                break;

            // 错误
            case Cell.CELL_TYPE_ERROR:
                strCellValue = "错误";
                break;

            // 公式
            case Cell.CELL_TYPE_FORMULA:
                strCellValue = "公式";
                break;

            default:
                strCellValue = cell.getCellComment().toString();
        }

        // 特殊字符处理
        /*strCellValue = getReplaceChar(strCellValue);*/

        return strCellValue;
    }

}
