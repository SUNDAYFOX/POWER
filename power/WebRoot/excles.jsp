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
	String trainingID=request.getParameter("trainingID");  
	String trainingTime=request.getParameter("trainingTime");  
	String period=request.getParameter("period");  
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html:html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title>培训班数据导入及导出</title>
	<link href="<%=path%>/css/portalet.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/font_color.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css">
	<script language="javascript" src="<%=path%>/js/jquery.js"></script>
	<script type="text/javascript" src="<%=path%>/js/util.js"></script>
	<script type="text/javascript">
		function init(d) {
			var str = encodeURI(encodeURI('<%=path%>/DB/template/校内开班通知.doc'));
			alert(str);
			//window.location.href=str;
			
		}
		function fileimport(str) {
			
			//openWindowCenter("searchPower.do?action=importOperate",800,600 );
			//alert(str);
			openWindowCenter("searchPower.do?action="+str+"&trainingID="+'<%=trainingID%>'+"&trainingTime="+'<%=trainingTime%>'+"&period="+'<%=period%>',800,600 );			
		}		
		function make(d) {
			//alert(d);
			//var res = "<%=path%>/searchPower.do?action=query";
            $.ajax({
                    url:"<%=path%>/makeData.do?action=query",
                    type:"Post",
                    data:{trainingID : '<%=trainingID%>'  },
                    dataType:"String",//jsonp 实现跨域
                    jsonpCallback:"fun",//服务器返回执行的方法名
                    success:function (data) {
                        //console.log("success");
                        alert("导入成功");
                    },
                    error:function (err) {
                        alert(err);
                    }
                });

         
			
		}
</script>	
</head>
<body >
<div id="Main">
  <div id="basic_info">
    <div class="infoTitle">
      <div class="left"><img src="<%=path%>/images/basic_info.gif" width="26" height="20" /> 培训班信息</div> 
      <div class="right"></div></div>
<div id="basic_infoMain">
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="15%" rowspan="23">
								<div id="logo_image">
									<div class="left">第一步：开发前期数据维护</div>

								</div>
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 培训课程表：</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('trainingcourseimport');" />
														
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 教师信息表：</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('teacherinfoimport');" />													
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 教材信息表：</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('bookinfoimport');" />
														
							</td>	
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 预算信息表：</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/fee.xls'" />
														
							</td>																					
						</tr>						
					</table>
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="15%" rowspan="23">
								<div id="logo_image">
									<div class="left">第二步：上级开班通知维护</div>

								</div>
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 上级开班通知：</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/superiorclass.doc'" />
														
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
							</td>
							<td width="10%">					
																			
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
							</td>
							<td width="10%">					
							
														
							</td>	
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
							</td>
							<td width="10%">					
							
														
							</td>	
																				
						</tr>						
					</table>		
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="15%" rowspan="23">
								<div id="logo_image">
									<div class="left">第三步：学员信息维护</div>

								</div>
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 学员信息维护：</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('studentinfoimport');" />
														
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
							</td>
							<td width="10%">					
																			
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
							</td>
							<td width="10%">					
							
														
							</td>	
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
							</td>
							<td width="10%">					
							
														
							</td>	
																				
						</tr>						
					</table>	
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td width="15%" rowspan="23">
									<div id="logo_image">
										<div class="left">第四步：运行数据维护</div>
	
									</div>
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 运行表维护：</b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('trainingrunimport');" />
															
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />校内开班通知： </b>
								</td>
								<td width="10%">					
									<input name="closePage" type="submit" class="Button" value="生成" onclick="make('schoolclass');" />											
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />学员手册： </b>
								</td>
								<td width="10%">					
								?
															
								</td>	
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />来客信息表： </b>
								</td>
								<td width="10%">					
								?
															
								</td>	
																					
							</tr>						
						</table>	
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td width="15%" rowspan="133">
									<div id="logo_image">
										<div class="left">第五步：运行后数据维护</div>
	
									</div>
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 学员收费明细表维护：</b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('chargeinfoimport');" />
															
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />教师报账表-维护： </b>
								</td>
								<td width="10%">					
																				
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />支出费用明细表： </b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('feeinfoimport');" />
															
								</td>	
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />来客信息表： </b>
								</td>
								<td width="10%">					
								
															
								</td>	
																					
							</tr>						
						</table>	
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td width="15%" rowspan="23">
									<div id="logo_image">
										<div class="left">第六步：一览表数据维护</div>
	
									</div>
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 一览表信息：</b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/superiorclass.doc'" />
															
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
								</td>
								<td width="10%">					
																				
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
								</td>
								<td width="10%">					
								
															
								</td>	
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> </b>
								</td>
								<td width="10%">					
								
															
								</td>	
																					
							</tr>						
						</table>																							
					<!-- <table width="80%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="20%" rowspan="23">
								<div id="logo_image">
									<div class="left">数据导入及导出</div>

								</div>
							</td>
							<td width="32%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 上级开班通知：</b>
							</td>
							<td>					
		
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/superiorclass.doc'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 校内开班通知：</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="生成" onclick="make('schoolclass');" />
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/result/schoolclass.doc'" />
														
							</td>
						</tr>
						<tr>
							<td>
							
							</td>
							<td>					
											
							</td>
						</tr>						
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 费用预算表：</b>
							</td>
							<td>					
		
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/fee.xls'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 教材信息表：</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('bookinfoimport');" />
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/bookinfo.xls'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 教师信息表：</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('teacherinfoimport');" />
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/teacherinfo.xls'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 培训计划表：</b>
							</td>
							<td>					
		
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/trainingplan.xls'" />
														
							</td>
						</tr>						
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 培训课程表：</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('trainingcourseimport');" />
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/trainingcourse.xls'" />
														
							</td>
						</tr>	
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> 培训运行表：</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="导入" onclick="fileimport('trainingrunimport');" />
							<input name="closePage" type="submit" class="Button" value="下载" onclick="javascript:window.location.href='<%=path%>/DB/template/trainingrun.xls'" />
														
							</td>
						</tr>								
					</table> -->					
					<input name="closePage" type="submit" class="Button" value="关闭" Onclick="window.close();" />
	</div>
	</div>
</div>
</body>
</html:html>
