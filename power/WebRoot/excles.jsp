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
	<title>��ѵ�����ݵ��뼰����</title>
	<link href="<%=path%>/css/portalet.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/font_color.css" rel="stylesheet" type="text/css">
	<link href="<%=path%>/css/publicStyle.css" rel="stylesheet" type="text/css">
	<script language="javascript" src="<%=path%>/js/jquery.js"></script>
	<script type="text/javascript" src="<%=path%>/js/util.js"></script>
	<script type="text/javascript">
		function init(d) {
			var str = encodeURI(encodeURI('<%=path%>/DB/template/У�ڿ���֪ͨ.doc'));
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
                    dataType:"String",//jsonp ʵ�ֿ���
                    jsonpCallback:"fun",//����������ִ�еķ�����
                    success:function (data) {
                        //console.log("success");
                        alert("����ɹ�");
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
      <div class="left"><img src="<%=path%>/images/basic_info.gif" width="26" height="20" /> ��ѵ����Ϣ</div> 
      <div class="right"></div></div>
<div id="basic_infoMain">
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="15%" rowspan="23">
								<div id="logo_image">
									<div class="left">��һ��������ǰ������ά��</div>

								</div>
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ��ѵ�γ̱�</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('trainingcourseimport');" />
														
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ��ʦ��Ϣ��</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('teacherinfoimport');" />													
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> �̲���Ϣ��</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('bookinfoimport');" />
														
							</td>	
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> Ԥ����Ϣ��</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/fee.xls'" />
														
							</td>																					
						</tr>						
					</table>
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<td width="15%" rowspan="23">
								<div id="logo_image">
									<div class="left">�ڶ������ϼ�����֪ͨά��</div>

								</div>
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> �ϼ�����֪ͨ��</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/superiorclass.doc'" />
														
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
									<div class="left">��������ѧԱ��Ϣά��</div>

								</div>
							</td>
							<td width="10%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ѧԱ��Ϣά����</b>
							</td>
							<td width="10%">					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('studentinfoimport');" />
														
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
										<div class="left">���Ĳ�����������ά��</div>
	
									</div>
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ���б�ά����</b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('trainingrunimport');" />
															
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />У�ڿ���֪ͨ�� </b>
								</td>
								<td width="10%">					
									<input name="closePage" type="submit" class="Button" value="����" onclick="make('schoolclass');" />											
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />ѧԱ�ֲ᣺ </b>
								</td>
								<td width="10%">					
								?
															
								</td>	
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />������Ϣ�� </b>
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
										<div class="left">���岽�����к�����ά��</div>
	
									</div>
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ѧԱ�շ���ϸ��ά����</b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('chargeinfoimport');" />
															
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />��ʦ���˱�-ά���� </b>
								</td>
								<td width="10%">					
																				
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />֧��������ϸ�� </b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('feeinfoimport');" />
															
								</td>	
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" />������Ϣ�� </b>
								</td>
								<td width="10%">					
								
															
								</td>	
																					
							</tr>						
						</table>	
					<table width="80%" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td width="15%" rowspan="23">
									<div id="logo_image">
										<div class="left">��������һ��������ά��</div>
	
									</div>
								</td>
								<td width="10%">
									<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> һ������Ϣ��</b>
								</td>
								<td width="10%">					
								<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/superiorclass.doc'" />
															
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
									<div class="left">���ݵ��뼰����</div>

								</div>
							</td>
							<td width="32%">
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> �ϼ�����֪ͨ��</b>
							</td>
							<td>					
		
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/superiorclass.doc'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> У�ڿ���֪ͨ��</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="����" onclick="make('schoolclass');" />
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/result/schoolclass.doc'" />
														
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
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ����Ԥ���</b>
							</td>
							<td>					
		
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/fee.xls'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> �̲���Ϣ��</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('bookinfoimport');" />
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/bookinfo.xls'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ��ʦ��Ϣ��</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('teacherinfoimport');" />
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/teacherinfo.xls'" />
														
							</td>
						</tr>
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ��ѵ�ƻ���</b>
							</td>
							<td>					
		
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/trainingplan.xls'" />
														
							</td>
						</tr>						
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ��ѵ�γ̱�</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('trainingcourseimport');" />
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/trainingcourse.xls'" />
														
							</td>
						</tr>	
						<tr>
							<td>
								<b class="font_cyan"><img src="<%=path%>/images/icon_left_1.gif" width="8" height="7" /> ��ѵ���б�</b>
							</td>
							<td>					
							<input name="closePage" type="submit" class="Button" value="����" onclick="fileimport('trainingrunimport');" />
							<input name="closePage" type="submit" class="Button" value="����" onclick="javascript:window.location.href='<%=path%>/DB/template/trainingrun.xls'" />
														
							</td>
						</tr>								
					</table> -->					
					<input name="closePage" type="submit" class="Button" value="�ر�" Onclick="window.close();" />
	</div>
	</div>
</div>
</body>
</html:html>
