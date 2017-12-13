<%@ page language="java" contentType="text/html; charset=gb2312"%>
<%@ taglib uri="http://struts.apache.org/tags-bean" prefix="bean" %>
<%@ taglib uri="http://struts.apache.org/tags-html" prefix="html" %>
<%@ taglib uri="http://struts.apache.org/tags-logic" prefix="logic" %>
<%@ taglib uri="http://struts.apache.org/tags-tiles" prefix="tiles" %>
<%
	String path = request.getContextPath();
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html:html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <title>收费信息表导入</title>
    <link href="<%=path%>/css/portalet.css" rel="stylesheet" type="text/css" />
	<link href="<%=path%>/css/font_color.css" rel="stylesheet" type="text/css" />
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css" />
	<SCRIPT language=JavaScript src="<%=path%>/js/calendar_popup.js"></script>
	<script type="text/javascript" src="<%=path%>/js/util.js"></script>
	<script type="text/javascript" src="<%=path%>/dwr/engine.js"> </script>
	<script type="text/javascript" src="<%=path%>/dwr/util.js"> </script>
	<script type="text/javascript" src="<%=path%>/dwr/interface/DwrTest.js"></script>
	<script language="javascript">
       function checkSubmit() {

				PowerForm.submit();			
	   }
	    
	</script>
  </head>
  
  <body>
  <html:form action="searchPower" method="post" enctype="multipart/form-data">
	<html:hidden property="action" value="chargeinfoimportOperate"/>
	<html:hidden property="trainingID"/>
	<html:hidden property="trainingTime"/>
	<html:hidden property="period"/>	
  <div id="Main">
  <div id="basic_info">
    <div class="infoTitle">
      <div class="left"><img src="<%=path%>/images/basic_info.gif" width="26" height="20" /> 收费信息导入：</div> 
      <div class="right"></div>
    </div>
	  <div id="basic_infoMain">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr>
				    	<td width="15%" align="right">
				    	收费信息表：<html:file property="uploadFile" />
					</tr>
					<tr>
					<td></td>
					</tr>
					<tr>
					<td colspan="2">
					<div align="center">
					<html:button property="button" styleClass="Button" value="导入"
						onclick="checkSubmit()" />					
					<html:button property="button" styleClass="Button" value="关闭"
						onclick="window.close();" />
					</div>
					</td>
					</tr>
				</table>
		  </div>
		 </div>
	</div>
		</html:form>
	
  </body>
</html:html>
