<%@ page language="java" contentType="text/html; charset=gb2312"%>
<%@ taglib uri="http://struts.apache.org/tags-bean" prefix="bean" %>
<%@ taglib uri="http://struts.apache.org/tags-html" prefix="html" %>
<%@ taglib uri="http://struts.apache.org/tags-logic" prefix="logic" %>
<%@ taglib uri="http://struts.apache.org/tags-tiles" prefix="tiles" %>
<%
	String path = request.getContextPath();
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <title>公司短期培训计划导入</title>
    <link href="<%=path%>/css/portalet.css" rel="stylesheet" type="text/css" />
	<link href="<%=path%>/css/font_color.css" rel="stylesheet" type="text/css" />
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css" />
	<SCRIPT language=JavaScript src="<%=path%>/js/calendar_popup.js"></script>
	<script type="text/javascript" src="<%=path%>/js/util.js"></script>
	<script type="text/javascript" src="<%=path%>/dwr/engine.js"> </script>
	<script type="text/javascript" src="<%=path%>/dwr/util.js"> </script>
	<script type="text/javascript" src="<%=path%>/dwr/interface/DwrTest.js"></script>
	<script language="javascript">

	    
	</script>
  </head>
  
  <body>
<form action="upload/upload" method="post" enctype ="multipart/form-data">
  <div id="Main">
  <div id="basic_info">
    <div class="infoTitle">
      <div class="left"><img src="<%=path%>/images/basic_info.gif" width="26" height="20" /> 公司短期培训计划导入：</div> 
      <div class="right"></div>
    </div>
	  <div id="basic_infoMain">
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr>

         <input type="file" name="fileUpload" label=“上传文件"/>

					</tr>
					<tr>
					<td></td>
					</tr>
					<tr>
					<td colspan="2">
					<div align="center">
					<input name="username" type="test" />
					<input name="closePage" type="submit" class="Button" value="导入" />			
					<html:button property="button" styleClass="Button" value="关闭"
						onclick="window.close();" />
					</div>
					</td>
					</tr>
				</table>
		  </div>
		 </div>
	</div>
   </form>
	
  </body>

