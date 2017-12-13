package power;

import java.io.IOException;
import java.io.PrintWriter;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.struts.action.ActionForm;
import org.apache.struts.action.ActionForward;
import org.apache.struts.action.ActionMapping;
import org.apache.struts.actions.DispatchAction;

import jxl.read.biff.BiffException;

public class MakeDataAction extends DispatchAction{
    private static final long serialVersionUID = 1L;
    public ActionForward query(ActionMapping mapping, ActionForm form,
            HttpServletRequest request, HttpServletResponse response) throws BiffException, IOException
    {
    	PowerForm theform = (PowerForm) form;
    	ToExcle excle = new ToExcle();
    	try {
			excle.schoolclass(theform,request);
		} catch (SQLException e) {
			e.printStackTrace();
		}
        return mapping.findForward("search");
    }    

 
}
