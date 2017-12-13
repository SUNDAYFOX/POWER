package power;

import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.struts.action.ActionForm;
import org.apache.struts.action.ActionForward;
import org.apache.struts.action.ActionMapping;
import org.apache.struts.actions.DispatchAction;

import jxl.read.biff.BiffException;
import util.ReturnInfo;

public class SearchPowerAction extends DispatchAction
{
	  
    public ActionForward query(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
		List infoList = new ArrayList<PowerVo>();
		try {
			infoList = new PowerDao().getClassList(theform,request);
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	
		//request.setAttribute("pageController", theform.getPageController());
		request.setAttribute("infoList",infoList);

        return mapping.findForward("search");
    }
    public ActionForward importOperate(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {

        return mapping.findForward("import");
    }
	public ActionForward trainingplanimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importtrainingplanExcel(theform, request.getSession());
		}catch (Exception e) {
			e.printStackTrace();
		} 
		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			return mapping.findForward("Fail");
		}
	}	    
    public ActionForward trainingrunimport(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {

        return mapping.findForward("trainingrunimport");
    }    
	//��ѵ���б������
	public ActionForward trainingrunimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importTrainingrunExcel(theform, request.getSession());
		} catch (Exception e) {
			e.printStackTrace();
		} 

		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			//request.setAttribute("info", returnInfo.getInfo());
			return mapping.findForward("Fail");
		}
	}
    public ActionForward trainingcourseimport(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
        return mapping.findForward("trainingcourseimport");
    }    
	//��ѵ�γ̱������
	public ActionForward trainingcourseimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importtrainingcourseExcel(theform, request.getSession());
		} catch (Exception e) {
			e.printStackTrace();
		} 
		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			return mapping.findForward("Fail");
		}
	}	
	
    public ActionForward teacherinfoimport(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
        return mapping.findForward("teacherinfoimport");
    }    
	//��ʦ��Ϣ�������
	public ActionForward teacherinfoimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importteacherinfoExcel(theform, request.getSession());
		}catch (Exception e) {
			e.printStackTrace();
		} 
		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			return mapping.findForward("Fail");
		}
	}	
	//�̲���Ϣ����
    public ActionForward bookinfoimport(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
        return mapping.findForward("bookinfoimport");
    }    
	//�̲���Ϣ�������
	public ActionForward bookinfoimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importbookinfoExcel(theform, request.getSession());
		}catch (Exception e) {
			e.printStackTrace();
		} 
		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			return mapping.findForward("Fail");
		}
	}		
	//ѧ Ա����
    public ActionForward studentinfoimport(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
        return mapping.findForward("studentinfoimport");
    }    
	//ѧ Ա�������
	public ActionForward studentinfoimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importStudentinfoExcel(theform, request.getSession());
		}catch (Exception e) {
			e.printStackTrace();
		} 
		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			return mapping.findForward("Fail");
		}
	}	
	//�շ���Ϣ����
    public ActionForward chargeinfoimport(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
        return mapping.findForward("chargeinfoimport");
    }    

	public ActionForward chargeinfoimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importChargeinfoExcel(theform, request.getSession());
		}catch (Exception e) {
			e.printStackTrace();
		} 
		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			return mapping.findForward("Fail");
		}
	}	
	//֧��������Ϣ����
    public ActionForward feeinfoimport(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
        return mapping.findForward("feeinfoimport");
    }    

	public ActionForward feeinfoimportOperate(ActionMapping mapping, ActionForm form,
			HttpServletRequest request, HttpServletResponse response)
	throws IOException, ServletException{
		PowerForm theform = (PowerForm) form;
		ReturnInfo returnInfo = new ReturnInfo();

		try {
			returnInfo = new PowerExcleDao().importFeeinfoExcel(theform, request.getSession());
		}catch (Exception e) {
			e.printStackTrace();
		} 
		if(returnInfo.isFlag()){
			return mapping.findForward("Success");
		}else{
			return mapping.findForward("Fail");
		}
	}	
}
