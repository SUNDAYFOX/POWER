package util;

import java.io.IOException;

import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;

/**
 * 
 * <p>Title: 字符集过滤器</p>
 * <p>Description: 字符集控制</p>
 * <p>Copyright: Copyright (c) 2007</p>
 * <p>Company: SI-TECH </p>
 * @author zhangli
 * @version 1.0
 *
 */

public class CharacterEncodingFilter 
    implements javax.servlet.Filter {
	
    protected FilterConfig filterConfig;
    protected String encodingName;
    protected boolean enable;

    /**
	 * 构造器
	 */
    public CharacterEncodingFilter() {
        this.encodingName = "gb2312";
        this.enable = false;
	}

    /**
     * @param filterConfig: 过滤器配置对象
	 * @throws ServletException
	 */
    public void init(FilterConfig filterConfig) 
        throws ServletException {
            this.filterConfig = filterConfig;
            loadConfigParams();
    }

    /**
	 * 载入配置项内容
	 */
    private void loadConfigParams() {
        this.encodingName = this.filterConfig.getInitParameter("encoding");
        String strIgnoreFlag = this.filterConfig.getInitParameter("enable");
        if (strIgnoreFlag.equalsIgnoreCase("true")) {
            this.enable = true;
        } else {
            this.enable = false;
        }
    }

    /**
     * @param request: servlet请求
     * @param response: servlet响应
     * @param chain: 过滤器
	 * @throws IOException,ervletException
	 */ 
	public void doFilter(ServletRequest request, ServletResponse response, FilterChain chain) 
        throws IOException, ServletException {
		    if (this.enable) {
			    request.setCharacterEncoding(this.encodingName);
		    }
		    chain.doFilter(request, response);
    }
	
	/**
     * 过滤器回收
	 */ 
    public void destroy() {
    }
}
