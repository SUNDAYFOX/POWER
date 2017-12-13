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
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>客户经理日志</title>
	<link href="<%=path%>/css/portalet.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/font_color.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/index.css" rel="stylesheet" />
	<link href="<%=path%>/css/win.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript">
		function checkSubmit() {
			PowerForm.action.value="searchPower";
			//alert(BankForm.action.value);
			PowerForm.submit();
		}
</script>
</head>

<body >
	<h1>power Example</h1>
	<script>
		window.location.href="searchPower.do?action=query";
	</script>
</body>
</html>