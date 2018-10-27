package top.evolutionary.excel.commons.context;

import top.evolutionary.excel.core.metaData.ExportParams;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.OutputStream;
import java.io.Serializable;

/**
 * @author richey
 */
public class Ihr360ExportExcelContext<T> implements Serializable {
    private static final long serialVersionUID = -879884203688640107L;

    private ExportParams<T> exportParams;

    private Workbook currentWorkbook;

    private Sheet currentSheet;

    private OutputStream outputStream;


    public ExportParams<T> getExportParams() {
        return exportParams;
    }

    public void setExportParams(ExportParams<T> exportParams) {
        this.exportParams = exportParams;
    }

    public Workbook getCurrentWorkbook() {
        return currentWorkbook;
    }

    public void setCurrentWorkbook(Workbook currentWorkbook) {
        this.currentWorkbook = currentWorkbook;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public void setOutputStream(OutputStream outputStream) {
        this.outputStream = outputStream;
    }

}
