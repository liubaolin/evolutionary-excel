package com.ihr360.excel.util;

import com.ihr360.excel.commons.context.Ihr360ExportExcelContext;
import com.ihr360.excel.commons.context.Ihr360ExportExcelContextHolder;
import com.ihr360.excel.commons.exception.ExcelException;
import com.ihr360.excel.core.cellstyle.ExcelCellStyle;
import com.ihr360.excel.core.metaData.ExcelSheet;
import com.ihr360.excel.core.metaData.ExportParams;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import static com.ihr360.excel.util.helper.Ihr360ExcelSheetHelper.write2Sheet;

public class Ihr360ExcelExportUtil {

    private static Logger logger = LoggerFactory.getLogger(Ihr360ExcelExportUtil.class);


    private Ihr360ExcelExportUtil() {

    }


    /**
     * 单个sheet导出
     * 导出数据可以是：
     * 1.javabean类型的对象集合
     * 2.Map类型的对象集合
     * 3.表头顺序由Map的key顺序决定
     * 4.如果导出javabean类型对象，数据顺序由@ExcellCell注解的index属性决定
     *
     * @param <T>
     * @param out 与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     */
    public static <T> void exportExcel(ExportParams<T> exportParams, OutputStream out) {
        try {
            Ihr360ExportExcelContextHolder.initExportContext(true, exportParams, out);
            exportExcel();
        } catch (Exception e) {
            logger.error(e.toString());
            throw new ExcelException(e);
        } finally {
            Ihr360ExportExcelContextHolder.clean();
        }
    }



    /**
     * //todo 暂未使用,待改造
     * 将一个二维数组导出，一维是行，二维是行的cell值
     * <p>
     *
     * @param datalist
     * @param out
     */
    public static void exportExcel(String[][] datalist, OutputStream out) {
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            // 生成一个表格
            HSSFSheet sheet = workbook.createSheet();

            for (int i = 0; i < datalist.length; i++) {
                String[] r = datalist[i];
                HSSFRow row = sheet.createRow(i);
                for (int j = 0; j < r.length; j++) {
                    HSSFCell cell = row.createCell(j);
                    //cell max length 32767
                    if (r[j].length() > 32767) {
                        r[j] = "--此字段过长(超过32767),已被截断--" + r[j];
                        r[j] = r[j].substring(0, 32766);
                    }
                    cell.setCellValue(r[j]);
                }
            }
            //自动列宽
            if (datalist.length > 0) {
                int colcount = datalist[0].length;
                for (int i = 0; i < colcount; i++) {
                    sheet.autoSizeColumn(i);
                }
            }
            workbook.write(out);
        } catch (IOException e) {
            logger.error(e.toString(), e);
            throw new ExcelException(e);
        }finally {
            Ihr360ExportExcelContextHolder.clean();
        }
    }

    private static <T> void writeWorkbook() {
        try {

            Ihr360ExportExcelContext<T> excelContext = Ihr360ExportExcelContextHolder.getExcelContext();
            Workbook workbook = excelContext.getCurrentWorkbook();
            workbook.write(excelContext.getOutputStream());
        } catch (Exception e) {
            logger.error(e.toString(), e);
        }
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符合一定条件的数据以EXCEL 的形式输出到指定IO设备上<br>
     * 用于多个sheet
     *
     * @param <T>
     */
    private static <T> void exportExcel() {

        Ihr360ExportExcelContext<T> excelContext = Ihr360ExportExcelContextHolder.getExcelContext();
        ExportParams<T> exportParams = excelContext.getExportParams();
        List<ExcelSheet<T>> sheets = exportParams.getSheets();

        if (CollectionUtils.isEmpty(sheets)) {
            write2Sheet();
        } else {
            Map<String, ExcelCellStyle> headerStyleMap = exportParams.getHeaderStyleMap();
            Workbook workbook = excelContext.getCurrentWorkbook();
            for (ExcelSheet<T> sheet : sheets) {
                // 生成一个表格
                Sheet sheetItem = workbook.createSheet(sheet.getSheetName());
                excelContext.setCurrentSheet(sheetItem);
                ExportParams<T> itemExportParams = new ExportParams<>();
                itemExportParams.setHeaderMap(sheet.getHeaders());
                itemExportParams.setRowDatas(sheet.getDataset());
                itemExportParams.setHeaderStyleMap(headerStyleMap);
                itemExportParams.setDataTypeMap(sheet.getDataTypeMap());
                excelContext.setExportParams(itemExportParams);
                write2Sheet();
            }
        }
        writeWorkbook();
    }




}
