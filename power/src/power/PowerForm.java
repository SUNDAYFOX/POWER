package power;

import org.apache.struts.action.ActionForm;
import org.apache.struts.upload.FormFile;



public class PowerForm extends ActionForm{
    //private PageController pageController = new PageController();
    
    protected String action = null;
    protected String method = null;
	private FormFile uploadFile;
	private String attachmentFilePath;
	private String trainingID;//培训自编号
	private String companyID;//企业编号
	private String projectID;//项目编号
	private String projectName;//项目编号
	private String period;
	public String getPeriod() {
		return period;
	}
	public void setPeriod(String period) {
		this.period = period;
	}
	public String getTrainingTime() {
		return trainingTime;
	}
	public void setTrainingTime(String trainingTime) {
		this.trainingTime = trainingTime;
	}
	public String getDept() {
		return dept;
	}
	public void setDept(String dept) {
		this.dept = dept;
	}
	public String getProjectWorkDutyPeople() {
		return projectWorkDutyPeople;
	}
	public void setProjectWorkDutyPeople(String projectWorkDutyPeople) {
		this.projectWorkDutyPeople = projectWorkDutyPeople;
	}
	private String trainingTime;
	private String dept;
	private String projectWorkDutyPeople;
	public String getAction() {
		return action;
	}
	public void setAction(String action) {
		this.action = action;
	}
	public String getMethod() {
		return method;
	}
	public void setMethod(String method) {
		this.method = method;
	}
	public FormFile getUploadFile() {
		return uploadFile;
	}
	public void setUploadFile(FormFile uploadFile) {
		this.uploadFile = uploadFile;
	}
	public String getAttachmentFilePath() {
		return attachmentFilePath;
	}
	public void setAttachmentFilePath(String attachmentFilePath) {
		this.attachmentFilePath = attachmentFilePath;
	}
	public String getTrainingID() {
		return trainingID;
	}
	public void setTrainingID(String trainingID) {
		this.trainingID = trainingID;
	}
	public String getCompanyID() {
		return companyID;
	}
	public void setCompanyID(String companyID) {
		this.companyID = companyID;
	}
	public String getProjectID() {
		return projectID;
	}
	public void setProjectID(String projectID) {
		this.projectID = projectID;
	}
	public String getProjectName() {
		return projectName;
	}
	public void setProjectName(String projectName) {
		this.projectName = projectName;
	}
	public String getTrainingTarget() {
		return trainingTarget;
	}
	public void setTrainingTarget(String trainingTarget) {
		this.trainingTarget = trainingTarget;
	}
	public String getDutyDept() {
		return dutyDept;
	}
	public void setDutyDept(String dutyDept) {
		this.dutyDept = dutyDept;
	}
	public String getDutyPeople() {
		return dutyPeople;
	}
	public void setDutyPeople(String dutyPeople) {
		this.dutyPeople = dutyPeople;
	}
	public String getDutyProjectPeople() {
		return dutyProjectPeople;
	}
	public void setDutyProjectPeople(String dutyProjectPeople) {
		this.dutyProjectPeople = dutyProjectPeople;
	}
	private String trainingTarget;//项目编号
	private String dutyDept;//责任部门
	private String dutyPeople;//责任部门
	private String dutyProjectPeople;//项目运行责任人
}
