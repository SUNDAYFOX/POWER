<%@ page language="java" contentType="text/html; charset=gb2312"%>
<%@ taglib uri="http://struts.apache.org/tags-bean" prefix="bean" %>
<%@ taglib uri="http://struts.apache.org/tags-html" prefix="html" %>
<%@ taglib uri="http://struts.apache.org/tags-logic" prefix="logic" %>
<%@ taglib uri="http://struts.apache.org/tags-tiles" prefix="tiles" %>
<%
	String funcCode = null;
	if (request.getParameter("funcCode") == null) {
	} else {
		funcCode = request.getParameter("funcCode");
		session.setAttribute("funcCode", funcCode);
	}
	String path = request.getContextPath();
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html:html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache">
	<meta http-equiv="expires" content="0">
	<title>操作成功</title>
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="succeed">
<div id="title">
   <div class="left">操作成功</div>
</div>
<br><br>
<table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="60" align="center"><img src="<%=path%>/images/popup_ok.gif" width="50" height="50" /></td>
    <td class="red">恭喜您，操作成功</td>
  </tr>
  <tr>
	<td colspan="2" >
		<div align="center" id="button">
		<input type="button" class="Button" value="关闭" onclick="window.close();"/>							
		</div>
	</td>
	</tr>
</table>
</div>

<script>
  	window.opener.location.reload();
</script>
</body>
</html:html>