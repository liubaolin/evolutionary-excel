package org.joyfulframework;

/**
 * 提供相关配置的常量
 * Created by richey on 16-9-15.
 * @author richey
 * @since 1.0.SNAPSHOT
 */
public interface ConfigConstant {
    String CONFIG_FILE = "joyful.properties";
    String JDBC_DRIVER = "joyful.framework.jdbc.driver";
    String JDBC_URL = "joyful.framework.jdbc.url";
    String JDBC_USERNAME = "joyful.framework.jdbc.username";
    String JDBC_PASSWORD = "joyful.framework.jdbc.password";

    String APP_BASE_PACKAGE = "joyful.framework.app.base_package";
    String APP_JSP_PATH = "joyful.framework.app.jap_path";
    String APP_ASSET_PATH = "joyful.framework.app.asset_path";

}
