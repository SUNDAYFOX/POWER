package util;



public class ReturnInfo {

    private String code = null;         //�����������
    private String info = null;         //���������Ϣ
    private boolean flag ;              //�������
    
    /**
     * ������
     */
    public ReturnInfo() {
        super();
    }

    /**
	 * @return ���ز����������
	 */
    public String getCode() {
        return code;
    }
    
    /**
     * @param code: �������
	 * @return ���ؿ�
	 */
    public void setCode(String code) {
        this.code = code;
    }
    
    /**
	 * @return ���ز������
	 */
    public boolean isFlag() {
        return flag;
    }
    
    /**
     * @param code: �������
	 * @return ���ؿ�
	 */
    public void setFlag(boolean flag) {
        this.flag = flag;
    }
    
    /**
	 * @return ���ز��������Ϣ
	 */
    public String getInfo() {
        return info;
    }
    
    /**
     * @param code: ���������Ϣ
	 * @return ���ؿ�
	 */
    public void setInfo(String info) {
        this.info = info;
    }
}
