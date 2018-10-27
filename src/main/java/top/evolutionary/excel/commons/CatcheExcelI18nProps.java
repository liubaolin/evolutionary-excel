package top.evolutionary.excel.commons;

import com.google.common.cache.Cache;
import com.google.common.cache.CacheBuilder;
import top.evolutionary.excel.commons.exception.ExcelException;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Properties;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.TimeUnit;

/**
 * @author richey
 */
public class CatcheExcelI18nProps {

    private static final Logger logger = LoggerFactory.getLogger(CatcheExcelI18nProps.class);
    private static final String I18N_COLUMN_SEPARATOR = ",";//配置文件多语言词条分割符


    private CatcheExcelI18nProps() {
        super();
    }

    private static Cache<String, Properties> propCatche = CacheBuilder.newBuilder()
            //设置cache的初始大小为10
            .initialCapacity(5)
            //设置缓存最大容量为100，超过100之后就会按照LRU最近虽少使用算法来移除缓存项
            .maximumSize(10)
            //缓存项在创建后，在24小时时间内没有被读/写访问，则清除。
            .expireAfterAccess(24, TimeUnit.HOURS)
            .build();

    public static Properties getProps(String path) throws ExecutionException {
        Properties properties;
        properties = propCatche.get(path, () -> {
            Properties newProps = new Properties();
            Resource resource = new ClassPathResource(path);
            try {
                InputStream inStream = resource.getInputStream();
                newProps.load(new InputStreamReader(inStream, "UTF-8"));
            } catch (IOException e) {
                logger.error(e.getMessage());
                throw new ExcelException(e);
            }
            return newProps;
        });
        return properties;
    }




    /**
     * 根据词条key获取所有多语言表头数据
     *
     * @param key
     * @return
     */
    private static String[] getPropHeaders(String key,String propsFileName) throws ExecutionException {
        if (StringUtils.isAnyBlank(key, propsFileName)) {
            return new String[0];

        }
        Properties properties = getProps(propsFileName);
        if (properties == null) {
            return new String[0];
        }

        String columnsStr = properties.getProperty(key);
        if (StringUtils.isBlank(columnsStr)) {
            throw new ExcelException("Excel i8n header key is missing: " + key);
        }
        return columnsStr.trim().split(I18N_COLUMN_SEPARATOR);
    }



    /**
     * 根据词条key获取根据多语言表头集合，顺序与词条文件多语言配置一致
     *
     * @param headerI18key
     * @return
     */
    public static List<String> getI18SortedHeaders(String headerI18key,String propsFileName) throws ExecutionException {
        if (StringUtils.isAnyBlank(headerI18key,propsFileName)) {
            return new ArrayList<>();
        }
        String[] headers = getPropHeaders(headerI18key,propsFileName);
        if (ArrayUtils.isEmpty(headers)) {
            return new ArrayList<>();
        }
        return Arrays.asList(headers);

    }
}
