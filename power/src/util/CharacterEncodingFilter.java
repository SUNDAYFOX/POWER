package util;

import java.io.IOException;

import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;

/**
 * 
 * <p>Title: �ַ���������</p>
 * <p>Description: �ַ�������</p>
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
	 * ������
	 */
    public CharacterEncodingFilter() {
        this.encodingName = "gb2312";
        this.enable = false;
	}

    /**
     * @param filterConfig: ���������ö���
	 * @throws ServletException
	 */
    public void init(FilterConfig filterConfig) 
        throws ServletException {
            this.filterConfig = filterConfig;
            loadConfigParams();
    }

    /**
	 * ��������������
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
     * @param request: servlet����
     * @param response: servlet��Ӧ
     * @param chain: ������
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
     * ����������
	 */ 
    public void destroy() {
    }
}
