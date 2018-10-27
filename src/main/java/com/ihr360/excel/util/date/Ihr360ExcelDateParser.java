package com.ihr360.excel.util.date;

import com.ihr360.excel.commons.exception.ExcelDateParseException;
import org.apache.commons.lang3.StringUtils;

import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * @author richey
 */
public final class Ihr360ExcelDateParser {

    private static final Map<String, String> DATE_FORMAT_REGEXPS = new ConcurrentHashMap<>();

    static {
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}:\\d{1,2}\\s(am|pm)\\s[a-z]{3},\\s(mon|tue|wed|thu|fri|sat|sun)\\s(january|february|march|april|may|june|july|august|september|october|november|december)\\s\\d{1,2},\\s\\d{4}$", "hh:mm a z, E MMMM dd, yyyy");
        DATE_FORMAT_REGEXPS.put("^\\d{8}$", "yyyyMMdd");
        DATE_FORMAT_REGEXPS.put("^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec).\\s\\d{1,2},\\s\\d{4}\\s\\d{1,2}:\\d{2}\\s(am|pm)\\s[a-z]{3}$", "MMM. dd, yyyy hh:mm a z");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}-\\d{1,2}-\\d{4}$", "dd-MM-yyyy");
        DATE_FORMAT_REGEXPS.put("^(january|february|march|april|may|june|july|august|september|october|november|december)\\s\\d{1,2},\\s\\d{4}\\s\\d{1,2}:\\d{2}\\s(am|pm)$", "MMMM dd, yyyy hh:mm a");
        DATE_FORMAT_REGEXPS.put("^\\d{4}-\\d{1,2}-\\d{1,2}$", "yyyy-MM-dd");
        DATE_FORMAT_REGEXPS.put("^\\d{4}-\\d{1,2}$", "yyyy-MM");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}/\\d{1,2}/\\d{4}$", "MM/dd/yyyy");
        DATE_FORMAT_REGEXPS.put("^\\d{4}/\\d{1,2}/\\d{1,2}$", "yyyy/MM/dd");
        DATE_FORMAT_REGEXPS.put("^\\[a-z]{2,},\\d{1,2}\\s\\d{4}$", "MMM dd, yyyy");
        DATE_FORMAT_REGEXPS.put("^(january|february|march|april|may|june|july|august|september|october|november|december)\\s\\d{1,2},\\s\\d{4}$", "MMMM dd, yyyy");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}\\s[a-z]{3}\\s\\d{4}$", "dd MMM yyyy");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}\\s[a-z]{4,}\\s\\d{4}$", "dd MMMM yyyy");
        DATE_FORMAT_REGEXPS.put("^\\d{12}$", "yyyyMMddHHmm");
        DATE_FORMAT_REGEXPS.put("^\\d{8}\\s\\d{4}$", "yyyyMMdd HHmm");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}-\\d{1,2}-\\d{4}\\s\\d{1,2}:\\d{2}$", "dd-MM-yyyy HH:mm");
        DATE_FORMAT_REGEXPS.put("^\\d{4}-\\d{1,2}-\\d{1,2}\\s\\d{1,2}:\\d{2}$", "yyyy-MM-dd HH:mm");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}/\\d{1,2}/\\d{4}\\s\\d{1,2}:\\d{2}$", "MM/dd/yyyy HH:mm");
        DATE_FORMAT_REGEXPS.put("^\\d{4}/\\d{1,2}/\\d{1,2}\\s\\d{1,2}:\\d{2}$", "yyyy/MM/dd HH:mm");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}:\\d{2}$", "HH:mm");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}\\s[a-z]{3}\\s\\d{4}\\s\\d{1,2}:\\d{2}$", "dd MMM yyyy HH:mm");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}\\s[a-z]{4,}\\s\\d{4}\\s\\d{1,2}:\\d{2}$", "dd MMMM yyyy HH:mm");
        DATE_FORMAT_REGEXPS.put("^\\d{14}$", "yyyyMMddHHmmss");
        DATE_FORMAT_REGEXPS.put("^\\d{8}\\s\\d{6}$", "yyyyMMdd HHmmss");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}-\\d{1,2}-\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "dd-MM-yyyy HH:mm:ss");
        DATE_FORMAT_REGEXPS.put("^\\d{4}-\\d{1,2}-\\d{1,2}\\s\\d{1,2}:\\d{2}:\\d{2}$", "yyyy-MM-dd HH:mm:ss");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}/\\d{1,2}/\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "MM/dd/yyyy HH:mm:ss");
        DATE_FORMAT_REGEXPS.put("^\\d{4}/\\d{1,2}/\\d{1,2}\\s\\d{1,2}:\\d{2}:\\d{2}$", "yyyy/MM/dd HH:mm:ss");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}\\s[a-z]{3}\\s\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "dd MMM yyyy HH:mm:ss");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}\\s[a-z]{4,}\\s\\d{4}\\s\\d{1,2}:\\d{2}:\\d{2}$", "dd MMMM yyyy HH:mm:ss");
        DATE_FORMAT_REGEXPS.put("^\\d{1,2}:\\d{2}:\\d{2}$", "HH:mm:ss");
    }

    /**
     * 自定义日期格式
     *
     * @param regex            格式正则
     * @param dateFormatString format String
     */
    public static void addRegexFormat(String regex, String dateFormatString) {
        DATE_FORMAT_REGEXPS.put(regex, dateFormatString);
    }

    /**
     * 字符串转Date，如果不支持则返回null
     *
     * @param dateStr
     * @return 成功返回timestamp，否则返回null
     * @throws ParseException
     */
    public static Date getDate(String dateStr) throws ParseException {
        if (StringUtils.isBlank(dateStr)) {
            return null;
        }
        String pattern = determineDateFormat(dateStr);
        if (StringUtils.isBlank(pattern)) {
            throw new ExcelDateParseException("Excel日期类型转换异常，不支持当前日期格式");
        }
        DateFormat dateFormat = new SimpleDateFormat(pattern);
        return dateFormat.parse(dateStr);

    }

    /**
     * 字符串转Timestamp，如果不支持则返回null
     *
     * @param dateStr
     * @return 成功返回timestamp，否则返回null
     * @throws ParseException
     */
    public static Timestamp getTimestamp(String dateStr) throws ParseException {
        if (StringUtils.isBlank(dateStr)) {
            return null;
        }
        String pattern = determineDateFormat(dateStr);
        if (StringUtils.isBlank(pattern)) {
            return null;
        }
        DateFormat dateFormat = new SimpleDateFormat(pattern);
        Date parsedDate = dateFormat.parse(dateStr);
        return new Timestamp(parsedDate.getTime());
    }


    /**
     * 根据给出的String找到与之匹配的pattern，如果未找到返回null
     *
     * @param dateStr string类型的日期
     * @return
     * @see SimpleDateFormat
     */
    public static String determineDateFormat(String dateStr) {
        if (StringUtils.isBlank(dateStr)) {
            return null;
        }
        for (String regExp : DATE_FORMAT_REGEXPS.keySet()) {
            if (dateStr.toLowerCase().matches(regExp)) {
                return DATE_FORMAT_REGEXPS.get(regExp);
            }
        }
        return null;
    }
}