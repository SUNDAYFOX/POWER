package util;



public class ReturnInfo {

    private String code = null;         //操作结果代码
    private String info = null;         //操作结果信息
    private boolean flag ;              //操作结果
    
    /**
     * 构造器
     */
    public ReturnInfo() {
        super();
    }

    /**
	 * @return 返回操作结果代码
	 */
    public String getCode() {
        return code;
    }
    
    /**
     * @param code: 操作结果
	 * @return 返回空
	 */
    public void setCode(String code) {
        this.code = code;
    }
    
    /**
	 * @return 返回操作结果
	 */
    public boolean isFlag() {
        return flag;
    }
    
    /**
     * @param code: 操作结果
	 * @return 返回空
	 */
    public void setFlag(boolean flag) {
        this.flag = flag;
    }
    
    /**
	 * @return 返回操作结果信息
	 */
    public String getInfo() {
        return info;
    }
    
    /**
     * @param code: 操作结果信息
	 * @return 返回空
	 */
    public void setInfo(String info) {
        this.info = info;
    }
}
