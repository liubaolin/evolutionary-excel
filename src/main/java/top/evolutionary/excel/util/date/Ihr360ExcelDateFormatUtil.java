package top.evolutionary.excel.util.date;

import org.apache.commons.lang3.time.FastDateFormat;

import javax.annotation.Nonnull;
import java.text.ParseException;
import java.util.Date;

/**
 * Date的parse()与format(), 采用Apache Common Lang中线程安全, 性能更佳的FastDateFormat
 *
 * 注意Common Lang版本，3.5版才使用StringBuilder，3.4及以前使用StringBuffer.
 *
 * 1. 常用格式的FastDateFormat定义
 *
 * 2. 日期格式不固定时的String<->Date 转换函数.
 *
 *
 * @author richey
 * @see FastDateFormat#parse(String)
 * @see FastDateFormat#format(Date)
 * @see FastDateFormat#format(long)
 *
 */
public class Ihr360ExcelDateFormatUtil {

    //以T分割日期和时间，并带时区信息，符合ISO8601规范
    public static final String PATTERN_ISO = "yyyy-MM-dd'T'HH:mm:ss.SSSSZZ";
    public static final String PATTERN_ISO_ON_SECOND = "yyyy-MM-dd'T'HH:mm:ssZZ";
    public static final String PATTERN_ISO_ON_DATE = "yyyy-MM-dd";

    //以空格分割日期和时间，不带时区信息
    public static final String PATTERN_DEFAULT = "yyyy-MM-dd HH:mm:ss.SSS";
    public static final String PATTERN_DEFAULT_ON_SECOND = "yyyy-MM-dd HH:mm:ss";

    public static final String PATTERN_DEFAULT_HMS = "HH:mm:ss";



    //使用工厂方法FastDateFormat.getInstance(),从缓存中获取实例

    //以T分割日期和时间，并带时区信息，符合ISO8601规范
    public static final FastDateFormat ISO_FORMAT = FastDateFormat.getInstance(PATTERN_ISO);
    public static final FastDateFormat ISO_FORMAT_ON_SECOND = FastDateFormat.getInstance(PATTERN_ISO_ON_SECOND);
    public static final FastDateFormat ISO_FORMAT_ON_DATE = FastDateFormat.getInstance(PATTERN_ISO_ON_DATE);

    //以空格分割日期和时间，不带时区信息
    public static final FastDateFormat DEFAULT_FORMAT = FastDateFormat.getInstance(PATTERN_DEFAULT);
    public static final FastDateFormat DEFAULT_ON_SECOND_FORMAT = FastDateFormat.getInstance(PATTERN_DEFAULT_ON_SECOND);
    public static final FastDateFormat ISO_FORMAT_HMS = FastDateFormat.getInstance(PATTERN_DEFAULT_HMS);

    /**
     * 解析日期字符串，仅用与pattern不固定的情况
     *
     * 否则直接使用本类中封装号的FastDateFormat
     *
     * FastDateFormat.getInstance()已经做了缓存，不会每次创建新对象，但直接使用对象仍能减少在缓存中的查找
     */
    public static Date parseDate(@Nonnull String pattern, @Nonnull String dateString) throws ParseException {
        return FastDateFormat.getInstance(pattern).parse(dateString);
    }

    /**
     * 解析日期字符串，仅用与pattern不固定的情况
     *
     * 否则直接使用本类中封装号的FastDateFormat
     *
     * FastDateFormat.getInstance()已经做了缓存，不会每次创建新对象，但直接使用对象仍能减少在缓存中的查找
     */
    public static String formatDate(@Nonnull String pattern, long date) {
        return FastDateFormat.getInstance(pattern).format(date);
    }
}