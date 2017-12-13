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
	<title>��ѵ�ƻ�</title>
	<link href="<%=path%>/css/portalet.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/font_color.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/index.css" rel="stylesheet" />
	<link href="<%=path%>/css/win.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="<%=path%>/js/util.js"></script>
	<script type="text/javascript">
		function checkSubmit() {
			PowerForm.action.value="query";
			
			PowerForm.submit();
		}
	function goExcle(trainingID,trainingTime,period){
		openWindowCenter("<%=path%>/excles.jsp?trainingID="+ trainingID+"&trainingTime="+trainingTime+"&period="+period,1200,550); 
	}		
	function do_save(){
		//window.location.href='bank.do?action=importExcel'
		openWindowCenter("searchPower.do?action=importOperate",800,600 );
	}	
	function do_save1(){
		//window.location.href='bank.do?action=importExcel'
		openWindowCenter("searchPower.do?action=importOperate",800,600 );
	}			
</script>
</head>

<body >
    
<html:form action="searchPower" method="post">    
<html:hidden property="action" value="query"/>   
<div id="Main">
	<div class="groupItem2">
		<div class="itemHeader">
	  	<div id="zzi">��ѵ�ƻ�</div>
	  		<div id="tu">
			  <div id="sub">
				 <div class="li"><a href="javascript:formMng();"><img src="<%=path%>/images/jian.gif" border="0" width="15" height="15"></a></div>
			  </div>
			</div>
		</div>
		
   	<div id="searchForm" class="itemContent" style="display:block">
		<div id="pop">
				<table class="table02" width="100%" border="0" cellpadding="3" cellspacing="0">
					<tr>
					<img src="<%=path%>/images/seach_arrow.gif" />&nbsp;��Ŀ����������
					<html:text property="projectWorkDutyPeople" maxlength="10"></html:text>   
		 			
		 			<input type="button" class="Button" value="��ѯ" onclick="checkSubmit()" />
				</table>
				
		<div id="table01">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">					
						<div id="button_area">
						
							<div class="btn1">
							<a href ="<%=path%>/DB/classanalysis.xls" target="_blank">���ڷ���</a></div>
							<div class="btn1">	
							<a href="javascript:do_save1();">��ѵ�ƻ�����</a></div>
							
						</div>			
						<tr class="grey">
						</tr>			
						<tr>
							<th width="8%">
								��ѵ������
							</th>
							<th width="8%">
								��ѵʱ��
							</th>
							<th width="8%">
								���β���
							</th>
							<th width="10%">
								��Ŀ����������
							</th>
							<th width="8%">
								��Ŀ����������
							</th>		
							<th width="8%">
								���ڴ�������
							</th>																				
							<th width="8%">
								��ϸ
							</th>							
						</tr>
						<logic:present name="infoList">
						<logic:iterate id="vo" name="infoList" type="power.PowerVo">

							<tr>
									<td > 
									&nbsp;<bean:write name="vo" property="className" /></a> 
									
								</td>
								<td>
									&nbsp;
									<bean:write name="vo" property="trainingTime" />��
								</td>		
								<td>
									&nbsp;
									<bean:write name="vo" property="dutyDept" />
								</td>	
																<td>
									&nbsp;
									<bean:write name="vo" property="projectDevelopmentDutyPeople" />
								</td>	
								<td>
									&nbsp;
									<bean:write name="vo" property="projectWorkDutyPeople" />
								</td>	
								<td>
									&nbsp;
									<bean:write name="vo" property="store" />
								</td>	
								<td>
									<div align="center">
										<img src="<%=path%>/images/button_seach.gif" style='cursor: hand' alt="����֪ͨ"
											onclick="javascript:goExcle('<bean:write name="vo" property="trainingID" />','<bean:write name="vo" property="trainingTime" />','<bean:write name="vo" property="period" />');">
									</div>
								</td>																										
							</tr>
						</logic:iterate>
					</logic:present>	
				</table>
  </div>
</div>
          </html:form>
</body>
</html>