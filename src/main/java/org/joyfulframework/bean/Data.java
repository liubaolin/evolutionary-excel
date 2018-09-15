package org.joyfulframework.bean;

/**
 * 返回数据对象
 * Created by richey on 16-9-15.
 * @author richey
 * @since 1.0.SNAPSHOT
 */
public class Data {
    /**
     * 数据模型
     */
    private Object model;
    public Data(Object model){
        this.model = model;
    }

    public Object getModel() {
        return model;
    }
}
