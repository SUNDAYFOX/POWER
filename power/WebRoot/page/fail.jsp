<%@ page language="java" contentType="text/html; charset=gb2312"%>
<%
	String path = request.getContextPath();
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html:html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache">
	<meta http-equiv="expires" content="0">
	<title>����ʧ����ʾ</title>
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css">
</head>
<body>
<div id="succeed">
<div id="title">
   <div class="left">����ʧ��</div>
</div>
<br><br>
<table width="90%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="60" align="center"><img src="<%=path%>/images/popup_x.gif" width="50" height="50" /></td>
    <td class="red">�Բ���<%=request.getAttribute("info")%></td>
  </tr>
  <tr>
	<td colspan="2" >
		<div align="center" id="button">
		<input type="button" class="Button" value="�ر�" onclick="window.close();"/>							
		</div>
	</td>
	</tr>
</table>
</div>

</body>
</html:html>