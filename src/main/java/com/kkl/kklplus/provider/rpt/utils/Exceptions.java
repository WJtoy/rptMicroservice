/**
 * Copyright (c) 2005-2012 springside.org.cn
 */
package com.kkl.kklplus.provider.rpt.utils;

import javax.servlet.http.HttpServletRequest;
import java.io.PrintWriter;
import java.io.StringWriter;

/**
 * 关于异常的工具类.
 *
 * @author calvin
 * @version 2013-01-15
 */
public class Exceptions {

    /**
     * 将CheckedException转换为UncheckedException.
     */
    public static RuntimeException unchecked(Exception e) {
        if (e instanceof RuntimeException) {
            return RuntimeException.class.cast(e);
        } else {
            return new RuntimeException(e);
        }
    }

    /**
     * 将ErrorStack转化为String.
     */
    public static String getStackTraceAsString(Throwable e) {
        if (e == null)
            return "";
        StringWriter stringWriter = new StringWriter();
        e.printStackTrace(new PrintWriter(stringWriter));
        return stringWriter.toString();
    }

    /**
     * 判断异常是否由某些底层的异常引起.
     */
    public static boolean isCausedBy(Exception ex, Class<? extends Exception>... causeExceptionClasses) {
        Throwable cause = ex.getCause();
        while (cause != null) {
            for (Class<? extends Exception> causeClass : causeExceptionClasses)
                if (causeClass.isInstance(cause))
                    return true;
            cause = cause.getCause();
        }
        return false;
    }

    /**
     * 获得最终的异常
     * getRootCause(ex).getMessage();
     */
    public static Throwable getRootCause(Throwable throwable) {
        if (throwable.getCause() != null)
            return getRootCause(throwable.getCause());

        return throwable;
    }

    /**
     * 在request中获取异常类
     *
     * @param request
     * @return
     */
    public static Throwable getThrowable(HttpServletRequest request) {
        Throwable ex = null;
        Object obj = null;
        try {
            if (request.getAttribute("javax.servlet.error.exception") != null) {
                obj = request.getAttribute("javax.servlet.error.exception");
            }else if (request.getAttribute("exception") != null) {
                obj = request.getAttribute("exception");
            }
            ex = Throwable.class.cast(obj);
        } catch (Exception e) {
            //e.printStackTrace();
            if(obj != null) {
                ex = new Throwable(obj.toString());
            }
        }
        return ex;
    }

}
