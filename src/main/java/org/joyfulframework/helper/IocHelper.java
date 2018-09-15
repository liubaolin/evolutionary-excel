package org.joyfulframework.helper;

import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.joyful4j.modules.utils.reflect.ReflectionUtil;
import org.joyfulframework.annotation.Inject;

import java.lang.reflect.Field;
import java.util.Map;

/**
 * 依赖注入助手类
 * Created by richey on 16-9-15.
 * @author richey
 * @since 1.0.SNAPSHOT
 */
public class IocHelper {
    static {
        //获取beanMap(内存所有bean与bean实例的对应关系)
        Map<Class<?>,Object> beanMap = BeanHelper.getBeanMap();
        if(MapUtils.isNotEmpty(beanMap)){
            //遍历beanMap
            for (Map.Entry<Class<?>,Object> beanEntry:beanMap.entrySet()) {
                Class<?> beanClass = beanEntry.getKey();
                Object beanInstance = beanEntry.getValue();
                //获取bean类定义的所有成员变量
                Field[] beanFiles = beanClass.getDeclaredFields();
                if(ArrayUtils.isNotEmpty(beanFiles)){
                    for (Field beanFile:beanFiles) {
                        if(beanFile.isAnnotationPresent(Inject.class)){
                            //在beanMap中获取beanFile对应的实例
                            Class<?> beanFileClass = beanFile.getType();
                            Object beanFileInstance = beanMap.get(beanFileClass);
                            if (beanFileInstance != null){
                                //通过反射初始话beanFile的值
                                ReflectionUtil.setField(beanInstance,beanFile,beanFileInstance);
                            }
                        }
                    }
                }
            }
        }
    }
}
