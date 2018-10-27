package com.ihr360.excel.commons.logs;

public enum ExcelLogType {

    CUSTOM_LOG_TYPE("excellog_type.enum.custom_log_type", "自定义Excel日志类型", "自定义Excel日志类型"),
    CUSTOM_LOG_POST_WARM("excellog_type.enum.custom_log_post_warm", "Excel标准导入后置提醒日志", "Excel标准导入后置提醒日志"),
    FIRST_ROW_RULE("excellog_type.enum.first_row_rule", "第一行不能隐藏或为空", "不能隐藏或为空"),
    NO_HEADER_ROW("excellog_type.enum.no_header_row", "没有满足条件的表头", "没有满足条件的表头"),
    HEADER_REQUIRED("excellog_type.enum.header_requied", "表头必须包含:{0}", "表头必须包含:{0}"),
    HIDDEN_ROW("excellog_type.enum.hidden_row", "第{0}行,是隐藏行", "该行是隐藏行"),
    BLANK_ROW("excellog_type.enum.hidden_row", "第{0}行,是空行", "该行是空行"),
    IGNORE_ROW("excellog_type.enum.ignore_row", "第{0}行,忽略未导入", "该行忽略未导入"),
    REQUIRED_COLUMN_HEADER_NOT_FOUND("excellog_type.enum.required_column_header_not_found", "未找到必填列{0}", "未找到必填列{0}"),
    UNSUPPORTED_TYPE("excellog_type.enum.unsupported_type", "不支持的类型{0}", "不支持的类型{0}"),
    COLUMN_DATA_REQUIRED("excellog_type.enum.column_data_required", "{0}列不能为空", "{0}列不能为空"),
    COLUMN_TYPE_CONSTRAINT("excellog_type.enum.column_type_constraint", "{0}类型只能是{1}", "{0}类型只能是{1}"),
    COLUMN_IN_SCOPE("excellog_type.enum.column_in_scope", "{0}取值范围只能是[{1}]", "{0}取值范围只能是[{1}]"),
    COLUMN_SCOPE_LT("excellog_type.enum.column_scope_lt", "{0}必须小于{1}", "{0}必须小于{1}"),
    COLUMN_SCOPE_GT("excellog_type.enum.column_scope_gt", "{0}必须大于{1}", "{0}必须大于{1}"),
    COLUMN_SCOPE_LE("excellog_type.enum.column_scope_le", "{0}必须小于等于{1}", "{0}必须小于等于{1}"),
    COLUMN_SCOPE_GE("excellog_type.enum.column_scope_ge", "{0}必须大于等于{1}", "{0}必须大于等于{1}"),
    COLUMN_CON_NOT_CONVERT_TO_DATE("excellog_type.enum.column_scope_ge", "{0}不是合法的日期类型", "{0}不是合法的日期类型"),
    COLUMN_FIELD_DATA_TYPE_ERR("excellog_type.enum.column_data_type_err", "{0}数据类型错误", "{0}数据类型错误"),
    ROW_COLUMN_FIELD_DATA_TYPE_ERR("excellog_type.enum.row_column_data_type_err", "第{0}行，{1}数据类型错误", "{1}数据类型错误"),
    EXCEL_COMMON_ENCRYPTED("excellog_type.enum.excel_common_encrypted", "Excel不能存在密码", "Excel不能存在密码"),
    EXCEL_COMMON_COMMON("excellog_type.enum.excel_common_common", "Excel格式错误", "Excel格式错误"),
    EXCEL_COMMON_FORMAT_ENCRYPTED("excellog_type.enum.excel_common_format_encrypted", "Excel格式错误或存在密码", "Excel格式错误或存在密码"),
    EXCEL_COMMON_NO_DATA("excellog_type.enum.excel_common_no_data", "Excel数据为空", "Excel数据为空");


    private String logKey;

    private String logMessage;

    /**
     * 不包含行号的错误信息，方便自定义日志展示格式
     */
    private String noLineNumberMsg;


    private ExcelLogType(String key, String logMessage, String noLineNumberMsg) {
        this.logKey = key;
        this.logMessage = logMessage;
        this.noLineNumberMsg = noLineNumberMsg;
    }

    public String getLogKey() {
        return logKey;
    }


    public String getLogMessage() {
        return logMessage;
    }

    public String getNoLineNumberMsg() {
        return noLineNumberMsg;
    }
}
