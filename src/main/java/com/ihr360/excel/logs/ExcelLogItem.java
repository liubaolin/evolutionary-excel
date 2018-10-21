package com.ihr360.excel.logs;

import org.springframework.util.StringUtils;

import java.text.MessageFormat;

public class ExcelLogItem {

    private ExcelLogType logType;

    private String logMsgKey;

    private String defaultLogMsg;

    protected Object[] args;

    private Integer colNum;

    public ExcelLogItem() {
    }

    public ExcelLogItem(String defaultLogMsg) {
        this.logType = ExcelLogType.CUSTOM_LOG_TYPE;
        this.defaultLogMsg = defaultLogMsg;
    }

    public ExcelLogItem(ExcelLogType logType, Object[] args) {
        this.logType = logType;
        this.args = args;
    }

    public ExcelLogItem(ExcelLogType logType, Object[] args, Integer columnIndex) {
        this.logType = logType;
        this.args = args;
        this.colNum = columnIndex;
    }

    public ExcelLogItem(String logMsgKey, Object[] args, String defaultLogMsg) {
        this.logMsgKey = logMsgKey;
        this.args = args;
        this.defaultLogMsg = defaultLogMsg;
    }

    public ExcelLogType getLogType() {
        return logType;
    }

    public void setLogType(ExcelLogType logType) {
        this.logType = logType;
    }

    public Object[] getArgs() {
        return args;
    }

    public void setArgs(Object[] args) {
        this.args = args;
    }

    public static ExcelLogItem createExcelItem(ExcelLogType logType, Object[] args, Integer columnIndex) {
        return new ExcelLogItem(logType, args, columnIndex);
    }

    public static ExcelLogItem createExcelItem(ExcelLogType logType, Object[] args) {
        return new ExcelLogItem(logType, args);
    }

    public static ExcelLogItem createExcelItem(String logMsgKey, Object[] args, String defaultLogMsg) {
        return new ExcelLogItem(logMsgKey, args, defaultLogMsg);
    }

    public String getLogMsgKey() {
        return logMsgKey;
    }

    public void setLogMsgKey(String logMsgKey) {
        this.logMsgKey = logMsgKey;
    }

    public String getDefaultLogMsg() {
        return defaultLogMsg;
    }

    public void setDefaultLogMsg(String defaultLogMsg) {
        this.defaultLogMsg = defaultLogMsg;
    }

    public String getMessage() {
        if (logType == ExcelLogType.CUSTOM_LOG_TYPE) {
            return defaultLogMsg;
        }
        return toString();
    }

    public Integer getColNum() {
        return colNum;
    }

    public void setColNum(Integer colNum) {
        this.colNum = colNum;
    }

    @Override
    public String toString() {

        if (this.logType == null || StringUtils.isEmpty(this.logType.getLogMessage())) {
            return this.getDefaultLogMsg();
        }

        return MessageFormat.format(this.logType.getLogMessage(), args);
    }

    public String getNoLineNumberMsg() {
        return MessageFormat.format(this.logType.getNoLineNumberMsg(), args);
    }

}
