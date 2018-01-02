package com.acmr.excel.interceptor;

import java.text.SimpleDateFormat;
import java.util.Date;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.springframework.web.servlet.HandlerInterceptor;
import org.springframework.web.servlet.ModelAndView;

import com.acmr.excel.controller.ExcelController;

public class LogInterceptor implements HandlerInterceptor {
	private static Logger log = Logger.getLogger(LogInterceptor.class);
	private SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmm");
	@Override
	public void afterCompletion(HttpServletRequest arg0,
			HttpServletResponse arg1, Object arg2, Exception arg3)
			throws Exception {
		// TODO Auto-generated method stub

	}

	@Override
	public void postHandle(HttpServletRequest arg0, HttpServletResponse arg1,
			Object arg2, ModelAndView arg3) throws Exception {

	}

	@Override
	public boolean preHandle(HttpServletRequest request,
			HttpServletResponse response, Object arg2) throws Exception {
		String ip = getIp2(request);
		String uri = request.getRequestURI();
		String excelId = request.getHeader("excelId");
		String step = request.getHeader("step");
		String date = sdf.format(new Date());
		String referer = request.getHeader("referer");
		String method = request.getMethod();
		String broser = request.getHeader("User-Agent");
		log.info(ip + "\t" + uri + "\t" + excelId + "\t" + step + "\t" + date +"\t" + referer + "\t"+method +"\t"+broser);
		return true;
	}

	public static String getIp2(HttpServletRequest request) {
        String ip = request.getHeader("X-Forwarded-For");
        if(StringUtils.isNotEmpty(ip) && !"unKnown".equalsIgnoreCase(ip)){
            //多次反向代理后会有多个ip值，第一个ip才是真实ip
            int index = ip.indexOf(",");
            if(index != -1){
                return ip.substring(0,index);
            }else{
                return ip;
            }
        }
        ip = request.getHeader("X-Real-IP");
        if(StringUtils.isNotEmpty(ip) && !"unKnown".equalsIgnoreCase(ip)){
            return ip;
        }
        return request.getRemoteAddr();
    }
}
