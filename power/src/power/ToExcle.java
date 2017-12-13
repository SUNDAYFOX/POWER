package power;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import javax.servlet.http.HttpServletRequest;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ToExcle {
	private boolean saveOnExit;  
    /** 
     * word文档 
     */  
    Dispatch doc = null;  
     
    /** 
     * word运行程序对象s 
     */  
    private   ActiveXComponent word;  
    /** 
     * 所有word文档 
     */  
    private   Dispatch documents;  
     
     
    /** 
     * 构造函数 
     */  
    public ToExcle() {  
        if(word==null){  
        word = new ActiveXComponent("Word.Application");  
        word.setProperty("Visible",new Variant(false));  
        }  
        if(documents==null)  
        documents = word.getProperty("Documents").toDispatch();  
        saveOnExit = false;  
    }  
     
    /** 
     * 设置参数：退出时是否保存 
     * @param saveOnExit boolean true-退出时保存文件，false-退出时不保存文件 
     */  
    public void setSaveOnExit(boolean saveOnExit) {  
        this.saveOnExit = saveOnExit;  
    }  
    /** 
     * 得到参数：退出时是否保存 
     * @return boolean true-退出时保存文件，false-退出时不保存文件 
     */  
    public boolean getSaveOnExit() {  
        return saveOnExit;  
    }  
     
    /** 
     * 打开文件 
     * @param inputDoc String 要打开的文件，全路径 
     * @return Dispatch 打开的文件 
     */  
    public Dispatch open(String inputDoc) {  
        return Dispatch.call(documents,"Open",inputDoc).toDispatch();  
    }  
     
    /** 
     * 选定内容 
     * @return Dispatch 选定的范围或插入点 
     */  
    public Dispatch select() {  
        return word.getProperty("Selection").toDispatch();  
    }  
     
    /** 
     * 把选定内容或插入点向上移动 
     * @param selection Dispatch 要移动的内容 
     * @param count int 移动的距离 
     */  
    public void moveUp(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++)  
            Dispatch.call(selection,"MoveUp");  
    }  
     
    /** 
     * 把选定内容或插入点向下移动 
     * @param selection Dispatch 要移动的内容 
     * @param count int 移动的距离 
     */  
    public void moveDown(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++)  
            Dispatch.call(selection,"MoveDown");  
    }  
     
    /** 
     * 把选定内容或插入点向左移动 
     * @param selection Dispatch 要移动的内容 
     * @param count int 移动的距离 
     */  
    public void moveLeft(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++) {  
            Dispatch.call(selection,"MoveLeft");  
        }  
    }  
     
    /** 
     * 把选定内容或插入点向右移动 
     * @param selection Dispatch 要移动的内容 
     * @param count int 移动的距离 
     */  
    public void moveRight(Dispatch selection,int count) {  
        for(int i = 0;i < count;i ++)  
            Dispatch.call(selection,"MoveRight");  
    }  
     
    /** 
     * 把插入点移动到文件首位置 
     * @param selection Dispatch 插入点 
     */  
    public void moveStart(Dispatch selection) {  
        Dispatch.call(selection,"HomeKey",new Variant(6));  
    }  
     
    /** 
     * 从选定内容或插入点开始查找文本 
     * @param selection Dispatch 选定内容 
     * @param toFindText String 要查找的文本 
     * @return boolean true-查找到并选中该文本，false-未查找到文本 
     */  
    public boolean find(Dispatch selection,String toFindText) {  
        //从selection所在位置开始查询  
        Dispatch find = word.call(selection,"Find").toDispatch();  
        //设置要查找的内容  
        Dispatch.put(find,"Text",toFindText);  
        //向前查找  
        Dispatch.put(find,"Forward","True");  
        //设置格式  
        Dispatch.put(find,"Format","True");  
        //大小写匹配  
        Dispatch.put(find,"MatchCase","True");  
        //全字匹配  
        Dispatch.put(find,"MatchWholeWord","True");  
        //查找并选中  
        return Dispatch.call(find,"Execute").getBoolean();  
    }  
     
    /** 
     * 把选定内容替换为设定文本 
     * @param selection Dispatch 选定内容 
     * @param newText String 替换为文本 
     */  
    public void replace(Dispatch selection,String newText) {  
        //设置替换文本  
        Dispatch.put(selection,"Text",newText);  
    }  
     
    /** 
     * 全局替换 
     * @param selection Dispatch 选定内容或起始插入点 
     * @param oldText String 要替换的文本 
     * @param newText String 替换为文本 
     */  
    public void replaceAll(Dispatch selection,String oldText,Object replaceObj) {  
        //移动到文件开头  
        moveStart(selection);  
         
        if(oldText.startsWith("table") || replaceObj instanceof ArrayList)  
            replaceTable(selection,oldText,(ArrayList) replaceObj);  
        else {  
            String newText = (String) replaceObj;  
            if(newText==null)  
                newText="";  
            if(oldText.indexOf("image") != -1&!newText.trim().equals("") || newText.lastIndexOf(".bmp") != -1 || newText.lastIndexOf(".jpg") != -1 || newText.lastIndexOf(".gif") != -1){  
                while(find(selection,oldText)) {  
                    replaceImage(selection,newText);  
                    Dispatch.call(selection,"MoveRight");  
                }  
            }else{  
                while(find(selection,oldText)) {  
                    replace(selection,newText);  
                    Dispatch.call(selection,"MoveRight");  
                }  
            }  
        }  
    }  
     
    /** 
     * 替换图片 
     * @param selection Dispatch 图片的插入点 
     * @param imagePath String 图片文件（全路径） 
     */  
    public void replaceImage(Dispatch selection,String imagePath) {  
        Dispatch.call(Dispatch.get(selection,"InLineShapes").toDispatch(),"AddPicture",imagePath);  
    }  
     
    /** 
     * 替换表格 
     * @param selection Dispatch 插入点 
     * @param tableName String 表格名称， 
     * 形如table$1@1、table$2@1...table$R@N，R代表从表格中的第N行开始填充，N代表word文件中的第N张表 
     * @param fields HashMap 表格中要替换的字段与数据的对应表 
     */  
    public void replaceTable(Dispatch selection,String tableName,ArrayList dataList) {  
        if(dataList.size() <= 1) {  
            System.out.println("Empty table!");  
            return;  
        }  
         
        //要填充的列  
        String[] cols = (String[]) dataList.get(0);  
         
        //表格序号  
        String tbIndex = tableName.substring(tableName.lastIndexOf("@") + 1);  
        //从第几行开始填充  
        int fromRow = Integer.parseInt(tableName.substring(tableName.lastIndexOf("$") + 1,tableName.lastIndexOf("@")));  
        //所有表格  
        Dispatch tables = Dispatch.get(doc,"Tables").toDispatch();  
        //要填充的表格  
        Dispatch table = Dispatch.call(tables,"Item",new Variant(tbIndex)).toDispatch();  
        //表格的所有行  
        Dispatch rows = Dispatch.get(table,"Rows").toDispatch();  
        //填充表格  
        for(int i = 1;i < dataList.size();i ++) {  
            //某一行数据  
            String[] datas = (String[]) dataList.get(i);  
             
            //在表格中添加一行  
            if(Dispatch.get(rows,"Count").getInt() < fromRow + i - 1)  
                Dispatch.call(rows,"Add");  
            //填充该行的相关列  
            for(int j = 0;j < datas.length;j ++) {  
                //得到单元格  
                Dispatch cell = Dispatch.call(table,"Cell",Integer.toString(fromRow + i - 1),cols[j]).toDispatch();  
                //选中单元格  
                Dispatch.call(cell,"Select");  
                //设置格式  
                Dispatch font = Dispatch.get(selection,"Font").toDispatch();  
                Dispatch.put(font,"Bold","0");  
                Dispatch.put(font,"Italic","0");  
                //输入数据  
                Dispatch.put(selection,"Text",datas[j]);  
            }  
        }  
    }  
     
    /** 
     * 保存文件 
     * @param outputPath String 输出文件（包含路径） 
     */  
    public void save(String outputPath) {  
        Dispatch.call(Dispatch.call(word,"WordBasic").getDispatch(),"FileSaveAs",outputPath);  
    }  
     
    /** 
     * 关闭文件 
     * @param document Dispatch 要关闭的文件 
     */  
    public void close(Dispatch doc) {  
        Dispatch.call(doc,"Close",new Variant(saveOnExit));  
        word.invoke("Quit",new Variant[]{});  
        word = null;  
    }  
     
    /** 
     * 根据模板、数据生成word文件 
     * @param inputPath String 模板文件（包含路径） 
     * @param outPath String 输出文件（包含路径） 
     * @param data HashMap 数据包（包含要填充的字段、对应的数据） 
     */  
    public void toWord(String inputPath,String outPath,HashMap data) {  
        String oldText;  
        Object newValue;  
        try {  
            if(doc==null)  
            doc = open(inputPath);  
             
            Dispatch selection = select();  
             
            Iterator keys = data.keySet().iterator();  
            while(keys.hasNext()) {  
                oldText = (String) keys.next();  
                newValue = data.get(oldText);  
                 
                replaceAll(selection,oldText,newValue);  
            }  
             
            save(outPath);  
        } catch(Exception e) {  
            System.out.println("操作word文件失败！");  
            e.printStackTrace();  
        } finally {  
            if(doc != null)  
                close(doc);  
        }  
    }  
     
    public synchronized static void word(String inputPath,String outPath,HashMap data){  
    	ToExcle j2w = new ToExcle();  
        j2w.toWord(inputPath,outPath,data);  
    }  
     
    @SuppressWarnings({ "rawtypes", "unchecked" })  
   public Boolean schoolclass(PowerForm theform,HttpServletRequest request) throws SQLException {  
    	Boolean res = true;
    	HashMap data = new HashMap();
	    try {
			Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  
	    String dbur1 = "jdbc:odbc:driver={Microsoft Access Driver (*.mdb)};DBQ=F://Database1.mdb";  
	    Connection conn;
		conn = DriverManager.getConnection(dbur1, "username", "password");
		Statement stmt = conn.createStatement();  
		Statement stmt2 = conn.createStatement();  
		String sql1 = ("select 培训自编号,项目名称,培训对象,项目开发部门,项目开发责任人,开发责任人电话,项目分口,项目运行责任人,运行责任人电话  from 短期培训计划表  where 培训自编号 = '"+theform.getTrainingID()+"'");  
	    String sql2 = ("select a.人数,a.开班时间,a.结束时间,a.培训地点,a.地点,a.班主任 ,b.办公电话, b.移动电话  from 培训班运行表  a left join 教师信息表 b on a.班主任=b.教师姓名"
	    		+ " where a.培训项目编号 = '"+theform.getTrainingID()+"'");  
	    ResultSet rs1 = stmt.executeQuery(sql1);  
	    ResultSet rs2 = stmt2.executeQuery(sql2);  
	    //stmt.execute(insertSql);
	    SimpleDateFormat format = new SimpleDateFormat("yyyy年MM月dd");
	    SimpleDateFormat format2 = new SimpleDateFormat("dd日");
    	data.put("$trainingid$","");
    	data.put("$classname$","");
    	data.put("$object$","");
    	data.put("$projectDevelopmentDept$","");
    	data.put("$projectDevelopmentPeople$","");
    	data.put("$projectPeoplePhone$","");
		data.put("$projectType$","");
		data.put("$verify$","");
		data.put("$projectworkPeople$","");
		data.put("$projectworkPeoplePhone$","");
    	data.put("$num$","");
    	data.put("$trainingtime$","");
    	data.put("$time$","");
    	data.put("$place$","");
    	data.put("$trainingplace$","");
    	data.put("$teacher1$","");
    	data.put("$teacher2$","");	   
    	data.put("$teacherphone1$","");
    	data.put("$teacherphone2$","");	   
	    while (rs1.next()) {  
	    	data.put("$trainingid$",rs1.getString(1));
	    	data.put("$classname$",rs1.getString(2));
	    	data.put("$object$",rs1.getString(3));
	    	data.put("$projectDevelopmentDept$",rs1.getString(4));
	    	data.put("$projectDevelopmentPeople$",rs1.getString(5));
	    	data.put("$projectPeoplePhone$",rs1.getString(6));
	    	String projectType = rs1.getString(7);
	    	if((projectType!=null&&projectType.equals("技术"))){
	    		data.put("$projectType$","技术技能培训部");
	    		data.put("$verify$","傅望");
	    	}else if(projectType!=null&&projectType.equals("管理")){
	    		data.put("$projectType$","管理培训部");
	    		data.put("$verify$","张银玲");
	    	}else{
	    		data.put("$projectType$","");
	    		data.put("$verify$","");	    		
	    	}
	    	data.put("$projectworkPeople$",rs1.getString(8));
	    	String projectworkPeoplePhone = rs1.getString(9);
	    	data.put("$projectworkPeoplePhone$",projectworkPeoplePhone.replace("/", "，"));
	    }  
	    while (rs2.next()) {  
	    	data.put("$num$",rs2.getString(1));
	    	Date startTime = rs2.getDate(2);
	    	Date endTime = rs2.getDate(3);
	    	String time = format.format(startTime)+"-"+format2.format(endTime);
	    	data.put("$trainingtime$",time);
	    	data.put("$time$",startTime+"日8：00前");
	    	String place = rs2.getString(4);
	    	if(place.equals("黄桷坪"))
	    	{
	    		place = "国网重庆技培中心本部（黄桷坪）招待所总台";
	    	}else if(place.equals("沙坪坝")){
	    		place = "国网重庆技培中心分部（沙坪坝）招待所总台";
	    	}else if(place.equals("四面山")){
	    		place = "国网重庆技培中心分部（四面山）招待所总台";
	    	}
	    	data.put("$place$",place);
	    	data.put("$trainingplace$",rs2.getString(5));
	    	String teacher = rs2.getString(6);
	    	String teacherPhone = rs2.getString(7);
	    	String teacherMobile = rs2.getString(8);
	    	if(teacher.indexOf("\\")>-1){
		    	data.put("$teacher1$",teacher.split("\\")[0]);
		    	data.put("$teacher2$",teacher.split("\\")[1]);	    		
	    	}else{
		    	data.put("$teacher1$",teacher);
		    	data.put("$teacher2$","");	
		    	if(teacherPhone == null){
		    		teacherPhone = "";
		    	}
		    	if(teacherMobile == null){
		    		teacherMobile = "";
		    	}
		    	data.put("$teacherphone1$",teacherPhone+"，"+teacherMobile);
		    	data.put("$teacherphone2$","");	   	
		    	
	    	}

	    	//data.put("$remark$",rs2.getString(6));
	    	//data.put("$classname$",rs1.getString(2));
	    	//data.put("$object$",rs1.getString(3));
	    }  
	    rs1.close();  
	    stmt.close();  
	    conn.close();            
        ToExcle j2w = new ToExcle();  
        long time1 = System.currentTimeMillis(); 
        System.out.println(request.getRealPath("")+"\\DB\\template\\校内开班通知.doc"); 
        ComThread.InitSTA();
        j2w.toWord(request.getRealPath("")+"\\DB\\template\\校内开班通知.doc",request.getRealPath("")+"\\DB\\result\\schoolclass.doc",data);  
        ComThread.Release(); 
        System.out.println("time cost : " + (System.currentTimeMillis() - time1));  
        
        return res;
    }  	

}
