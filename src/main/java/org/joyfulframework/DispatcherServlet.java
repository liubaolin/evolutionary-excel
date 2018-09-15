package org.joyfulframework;

import com.alibaba.fastjson.JSON;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.joyful4j.modules.utils.reflect.ReflectionUtil;
import org.joyfulframework.bean.Data;
import org.joyfulframework.bean.Handler;
import org.joyfulframework.bean.Param;
import org.joyfulframework.bean.View;
import org.joyfulframework.helper.BeanHelper;
import org.joyfulframework.helper.ConfigHelper;
import org.joyfulframework.helper.ControllerHelper;
import org.joyful4j.modules.utils.common.CodeUtil;
import org.joyful4j.modules.utils.common.StreamUtil;
import javax.servlet.ServletConfig;
import javax.servlet.ServletContext;
import javax.servlet.ServletException;
import javax.servlet.ServletRegistration;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.lang.reflect.Method;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;

/**
 * 请求转发器
 * Created by richey on 16-9-15.
 * @author richey
 * @since 1.0.SNAPSHOT
 */
@WebServlet(urlPatterns = "/*",loadOnStartup = 0)
public class DispatcherServlet extends HttpServlet{
    @Override
    public void init(ServletConfig servletConfig) throws ServletException {
        //初始化相关Helper类
        HelperLoader.init();
        //获取ServletContext对象
        ServletContext servletContext = servletConfig.getServletContext();
        //注册jsp的Servle
        ServletRegistration jspServlet = servletContext.getServletRegistration("jsp");
        jspServlet.addMapping(ConfigHelper.getAppJspPath()+"*");
        //注册处理静态资源的默认Servlet
        ServletRegistration defaultServlet = servletContext.getServletRegistration("default");
        defaultServlet.addMapping(ConfigHelper.getAppAssetPath()+"*");
    }

    @Override
    protected void service(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
        //获取请求方法与请求路径
        String requestMethod = req.getMethod().toLowerCase();
        String requestPath = req.getPathInfo();
        //获取Action处理器
        Handler handler = ControllerHelper.getHandler(requestMethod,requestPath);
        if(handler != null){
            Class<?> controllerClass = handler.getControllerClass();
            Object controllerBean = BeanHelper.getBean(controllerClass);
            //创建年请求参数对象
            Map<String,Object> parammMap = new HashMap<String,Object>();
            Enumeration<String> paramNames = req.getParameterNames();
            while (paramNames.hasMoreElements()){
                String paramName = paramNames.nextElement();
                String paramValue = req.getParameter(paramName);
                parammMap.put(paramName,paramValue);
            }
            String body = CodeUtil.decodeURL(StreamUtil.getString(req.getInputStream()));
            if(StringUtils.isNotBlank(body)){
                String[] params = body.split("&");
                if (ArrayUtils.isNotEmpty(params)) {
                    for (String param : params) {
                        String[] array = param.split("=");
                        if(ArrayUtils.isNotEmpty(array)&&array.length==2){
                            String paramName = array[0];
                            String paramValue = array[1];
                            parammMap.put(paramName,paramValue);
                        }
                    }

                }
            }
            Param param = new Param(parammMap);
            //调用Action方法
            Method actionMethod = handler.getMethod();
            Object result = ReflectionUtil.InvokeMethod(controllerBean,actionMethod,param);
            //处理Action方法的返回值
            if(result instanceof View){//返回jsp页面
                View view = (View) result;
                String path = view.getPath();
                if(StringUtils.isNotBlank(path)){
                    if(path.startsWith("/")){
                        resp.sendRedirect(req.getContextPath()+path);
                    }else{
                        Map<String,Object> model = view.getModel();
                        for (Map.Entry<String, Object> entry : model.entrySet()) {
                            req.setAttribute(entry.getKey(),entry.getValue());
                        }
                        req.getRequestDispatcher(ConfigHelper.getAppJspPath()+path).forward(req,resp);
                    }
                }
            }else if(result instanceof Data){//返回JSON数据
                Data data = (Data) result;
                Object model = data.getModel();
                if(model != null){
                    resp.setContentType("application/json");
                    resp.setCharacterEncoding("UTF-8");
                    PrintWriter writer = resp.getWriter();
                    String json = JSON.toJSONString(model);
                    writer.write(json);
                    writer.flush();
                    writer.close();
                }

            }
        }
    }
}
