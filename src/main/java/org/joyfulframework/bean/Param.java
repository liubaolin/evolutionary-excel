package org.joyfulframework.bean;

import org.joyful4j.modules.utils.common.CastUtil;

import java.util.Map;

/**
 * 请求参数对象
 * Created by richey on 16-9-15.
 * @author richey
 * @since 1.0.SNAPSHOT
 */
public class Param {
    private Map<String,Object> paramMap;

    public Param(Map<String,Object> paramMap){
        this.paramMap = paramMap;
    }

    /**
     * 根据参数名获取long类型数据
     * @param name
     * @return
     */
    public long getLong(String name){
        return CastUtil.castLong(paramMap.get(name));
    }

    /**
     * 获取所有字段信息
     * @return
     */
    public Map<String,Object> getParamMap(){
        return paramMap;
    }
}
