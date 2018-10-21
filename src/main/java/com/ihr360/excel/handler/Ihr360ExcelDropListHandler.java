package com.ihr360.excel.handler;

import com.ihr360.excel.metaData.ExportParams;
import org.apache.commons.collections.MapUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 下拉选处理类
 * @author richey
 */
public class Ihr360ExcelDropListHandler {


    /**
     * @param wb               HSSFWorkbook对象
     * @param realSheet        需要操作的sheet对象
     * @param datas            下拉的列表数据
     * @param startRow         开始行
     * @param endRow           结束行
     * @param startCol         开始列
     * @param endCol           结束列
     * @param hiddenSheetName  隐藏的sheet名
     * @param hiddenSheetIndex 隐藏的sheet索引 考虑到有多个列绑定下拉列表
     * @return
     * @throws Exception
     */
    public static HSSFWorkbook dropDownList2003(Workbook wb, Sheet realSheet, String[] datas, int startRow, int endRow,
                                                int startCol, int endCol, String hiddenSheetName, int hiddenSheetIndex) {

        HSSFWorkbook workbook = (HSSFWorkbook) wb;
        // 创建一个数据源sheet
        HSSFSheet hidden = workbook.createSheet(hiddenSheetName);
        // 数据源sheet页不显示
        workbook.setSheetHidden(hiddenSheetIndex, true);
        // 将下拉列表的数据放在数据源sheet上
        HSSFRow row = null;
        HSSFCell cell = null;
        for (int i = 0, length = datas.length; i < length; i++) {
            row = hidden.createRow(i);
            cell = row.createCell(0);
            cell.setCellValue(datas[i]);
        }
        DVConstraint constraint = DVConstraint.createFormulaListConstraint(hiddenSheetName + "!$A$1:$A" + datas.length);
        CellRangeAddressList addressList = null;
        HSSFDataValidation validation = null;
        row = null;
        cell = null;
        // 单元格样式
        CellStyle style = workbook.createCellStyle();
        style.setDataFormat(HSSFDataFormat.getBuiltinFormat("0"));
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        // 循环指定单元格下拉数据
        for (int i = startRow; i <= endRow; i++) {
            row = (HSSFRow) realSheet.createRow(i);
            cell = row.createCell(startCol);
            cell.setCellStyle(style);
            addressList = new CellRangeAddressList(i, i, startCol, endCol);
            validation = new HSSFDataValidation(addressList, constraint);
            realSheet.addValidationData(validation);
        }

        return workbook;
    }

    public static <T> Workbook handleDropDownList(ExportParams<T> exportParams, Workbook workbook, Sheet sheet, int startRow, int endRow) {
        Map<String, List<String>> dropDownsMap = exportParams.getDropDownsMap();
        Map<String, String> headers = exportParams.getHeaderMap();
        if (MapUtils.isNotEmpty(dropDownsMap)) {
            //linkHashMap的有序表头
            Set<String> headerKeys = headers.keySet();
            List<String> headerKeyList = new ArrayList<>(headerKeys);
            Set<String> dropHeaderKeys = dropDownsMap.keySet();
            int i = 1;
            for (String headerKey : dropHeaderKeys) {
                int columnIndex = headerKeyList.indexOf(headerKey);
                if (columnIndex < 0) {
                    continue;
                }
                List<String> dropList = dropDownsMap.get(headerKey);
                if (CollectionUtils.isEmpty(dropList)) {
                    continue;
                }
                String[] dropArray = new String[dropList.size()];

                for (int j = 0; j < dropList.size(); j++) {
                    dropArray[j] = dropList.get(j);
                }

                String hiddenSheetName = String.join("_", "hidden_sheet", i + "");
                workbook = dropDownList2003(workbook, sheet, dropArray, startRow, endRow, columnIndex, columnIndex, hiddenSheetName, i);
                i++;
            }
        }
        return workbook;
    }



}
