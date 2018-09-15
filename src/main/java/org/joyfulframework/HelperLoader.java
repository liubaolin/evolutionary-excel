package org.joyfulframework;

import org.joyfulframework.helper.BeanHelper;
import org.joyfulframework.helper.ClassHelper;
import org.joyfulframework.helper.ConfigHelper;
import org.joyfulframework.helper.IocHelper;
import org.joyful4j.modules.utils.common.ClassUtil;

/**
 * 集中加载相应的Helper类
 * Created by richey on 16-9-15.
 * @author richey
 * @since 1.0.SNAPSHOT
 */
public final class HelperLoader {
    public static void init(){
        Class<?>[] classList = {
                ClassHelper.class,
                BeanHelper.class,
                IocHelper.class,
                ConfigHelper.class
        };
        for (Class<?> cls : classList) {
            ClassUtil.loadClass(cls.getName(),true);
        }
    }
}
