package com.eiis5.word.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.eiis5.rulesmanage.utils.RulesManageUtils;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ImportWord {

	public static List<String> dagan = null;
	public static List<String> title = null;
	public static List<String> quan = null;
	public static List<String> contentAll = null;
	public static List<String> listStringName = null;
	public static List<Integer> outlineClass = null;
	public static List<String> serialNumber = null;
	public static List<Integer> nullSubscript = null;
	public static List<Integer> countXuhao = null;
	//
	public static List<String> outlineSerialNumber = null;
	public static List<String> nei = null;
	public static List<String> nei1 = null;
	public static List<String> outlineAll = null;
	public static int lenAll = 0;
	
	public static List<String> nn = null;
	
//	private static final Log log = LogFactory.getLog(ImportWord.class);
	public static String BATCHID = "";
	
	public static void main(String[] args) {
		try {
			readBookMarksByWordOrWps("E:\\测试\\文档1\\12.doc");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	//总方法
	public static void process(String filepath,String domainname,String guichengbm){
		BATCHID = getBatchId(domainname,"eiis.SEQ_EIIS_IMPORTWORD");
		moldAll(filepath,guichengbm,domainname);
		contentMethod(dagan,quan,guichengbm,domainname,filepath);
		updateBatchid(domainname,BATCHID);
	}
	
	//大纲的总管理
	public static void moldAll(String file,String connector,String domainname) {
		//获取大纲和全文
		try {
			wipeOutShowRevisions(file);
			readBookMarksByWordOrWps(file);
		} catch (IOException e) {
			// TODO Auto-generated catch block
//			log.error("异常："+e);
			
		}
		
		int i = dagan.size();
		dagan.remove(i-1);
		//获取
		serial(dagan);
		jumpClassFall();
		dislodgeListName(file);
		addOutline(file,connector,domainname);
	}
	
	
	//大纲下内容总管理
	public static void contentMethod(List<String> dagan,List<String> quan,String connector,String domainname,String filepath){
		dagan.add("~~~~~~~~~~~~~~");
		wordT(dagan,quan);
		try {
			addT004Content(domainname,connector,filepath);
		} catch (IOException e) {
			// TODO Auto-generated catch block
//			log.error("异常："+e);
		}
	}
	
	public static List<String> picture(String file) throws IOException{
		List<String> site = new ArrayList<String>();
		//hwpfDocument是专门处理word的， 在poi中还有处理其他office文档的类
		   HWPFDocument doc = new HWPFDocument(new FileInputStream(file));
		   //看看此文档有多少个段落
		   Range range = doc.getRange();
		   int numP = range.numParagraphs();
		   long start2 = System.currentTimeMillis();
		   //得到word的数据流
		   byte[] dataStream = doc.getDataStream();
		   int numChar = range.numCharacterRuns();
		   PicturesTable pTable = new PicturesTable(doc, dataStream, dataStream);
		   for(int j = 0; j < numChar; ++j){
			    CharacterRun cRun = range.getCharacterRun(j);    
			    //看看有没有图片
			    boolean has = pTable.hasPicture(cRun);
			    
			    if(has){
			    	int k = 0;
			    	while(cRun.text().replace("", "").trim().equals("")){
			    		CharacterRun cRun1 = range.getCharacterRun(j+k);
			    		
			    		if (cRun1.text().replace("", "").trim().length()!=0){
			    			break;
			    		}
			    		k++;
			    	}
			    		CharacterRun cRun2 = range.getCharacterRun(j+k);
			    		site.add(cRun2.text().replace("", "").trim());
			    Picture zhou = pTable.extractPicture(cRun, true);
			    }
		   }
		return site;
	}
	
	public static void dislodgePicture(String file) throws IOException{
		List<String> len = picture(file);
		for (int ii = 0; ii < len.size(); ii++) {
			for (int j = 0; j < quan.size(); j++) {
				if (quan.get(j).replaceAll("\\\r|\\\f", "").equals(len.get(ii).replaceAll("\\\r|\\\f", ""))) {
					quan.remove(j);
					contentAll.remove(j);
				}
			}	
		}
	}
	
	public static void jumpClassFall(){
		List<Integer> len = new ArrayList<Integer>();
		List<Integer> jump = new ArrayList<Integer>();
		for (int i = 0; i < outlineClass.size(); i++) {
			if (i==outlineClass.size()-1) {
				break;
			}
			switch (outlineClass.get(i)) {
			case 1:
				
				if (outlineClass.get(i+1)==1||outlineClass.get(i+1)==2 || outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+2);
				}
				break;
			case 2:
				if (outlineClass.get(i+1)==2||outlineClass.get(i+1)==3|| outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+2);
				}
				break;
			case 3:
				if (outlineClass.get(i+1)==3||outlineClass.get(i+1)==4|| outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+2);
				}
				break;
			case 4:
				if (outlineClass.get(i+1)==4||outlineClass.get(i+1)==5|| outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+2);
				}
				break;
			case 5:
				if (outlineClass.get(i+1)==5||outlineClass.get(i+1)==6|| outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+2);
				}
				break;
			case 6:
				if (outlineClass.get(i+1)==6||outlineClass.get(i+1)==7|| outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+2);
				}
				break;
			case 7:
				if (outlineClass.get(i+1)==7||outlineClass.get(i+1)==8|| outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+2);
				}
				break;
			case 8:
				if (outlineClass.get(i+1)==8||outlineClass.get(i+1)==9|| outlineClass.get(i)>outlineClass.get(i+1)) {
					
				}else{
					jump.add(outlineClass.get(i+1));
					outlineClass.remove(i+1);
					outlineClass.add(i+1, outlineClass.get(i)+1);
					len.add(i+1);
				}
				break;
			case 9:
				break;
			}
		}
		
		for (int j = 0; j < len.size(); j++) {
			for (int j2 = len.get(j); j2 < outlineClass.size(); j2++) {

					if (outlineClass.get(j2)==jump.get(j)) {
						outlineClass.remove(j2);
						outlineClass.add(j2, outlineClass.get(len.get(j)-2)+1);
					}else{
						break;
					}
				
			}
		}
	}
	
	public static void wipeOutShowRevisions(String filePath) throws IOException {		 	
	 	Dispatch wordFile = null;

		ActiveXComponent word = null;
	    try{
        word=new ActiveXComponent("Word.Application");  
          
        word.setProperty("Visible", new Variant(false));   
        //Dispatch:调度处理类，封装了一些操作来操作office，里面所有的可操作对象基本都是这种类型
          //获得所有文档对象
        Dispatch documents=word.getProperty("Documents").toDispatch();  
      //docName要打开的文档的详细地址
        wordFile=Dispatch.invoke(documents, "Open", Dispatch.Method, new Object[]{filePath,new Variant(true),new Variant(false)}, new int[1]).toDispatch();   
        Dispatch.put(wordFile,"TrackRevisions",new Variant(false));
        Dispatch.put(wordFile,"PrintRevisions",new Variant(false));
        Dispatch.put(wordFile,"ShowRevisions",new Variant(false));
		Dispatch.call((Dispatch) Dispatch.call(word, "WordBasic").getDispatch(),"FileSaveAs", filePath);
		/**另存为*/
		Dispatch.invoke(wordFile, "SaveAs", Dispatch.Method, new Object[] {    
				filePath, new Variant(true) }, new int[1]);
	    }finally{
        	try{
        		if (wordFile != null) {					
        			Dispatch.call(wordFile, "Close", false);
				}
        		if (word != null){
        			word.invoke("Quit", new Variant[] {});
        		}   
        	} catch (Exception e2) {
        		
        	}
	    }
	    //如果没有这句话，winword.exe进程将不会关闭
	   
	}

	//获取大纲和全文
	public static void readBookMarksByWordOrWps(String filePath) throws IOException {
		//用来存（判断前言和范围在一个级别）的值
		int llll = 0;
		dagan = new ArrayList<String>();
		title = new ArrayList<String>();
	 	quan = new ArrayList<String>();
	 	contentAll = new ArrayList<String>();
	 	listStringName = new ArrayList<String>();
	 	outlineClass = new ArrayList<Integer>();
	 	
	 	
	 	Dispatch wordFile = null;

		ActiveXComponent word = null;
	    try{
        word=new ActiveXComponent("Word.Application");  
          
        word.setProperty("Visible", new Variant(false));   
        //Dispatch:调度处理类，封装了一些操作来操作office，里面所有的可操作对象基本都是这种类型
          //获得所有文档对象
        Dispatch documents=word.getProperty("Documents").toDispatch();  
      //docName要打开的文档的详细地址
        wordFile=Dispatch.invoke(documents, "Open", Dispatch.Method, new Object[]{filePath,new Variant(true),new Variant(false)}, new int[1]).toDispatch();   
        Dispatch.put(wordFile, "ShowRevisions", false);
        //所有表格
        Dispatch tables = Dispatch.get(wordFile, "Tables").toDispatch(); 
        //获取表格总数
        int tableCount = Dispatch.get(tables, "Count").getInt();
        Dispatch table = null;
        //删除所有表格（删除第一个表格后，第二个表格会变成第一表格）
        for (int i = 0 ; i < tableCount ; i++) {
            table = Dispatch.call(tables, "Item", new Variant(1)).toDispatch();
            Dispatch.call(table, "Delete");
        }
        //所有段落
        Dispatch paragraphs=Dispatch.get(wordFile, "Paragraphs").toDispatch();  
          //段落总数
        int paraCount=Dispatch.get(paragraphs, "Count").getInt();

        int k=0;
//        int paraCount1 = 0;
          
       for(int i=0;i<paraCount;++i){  
    	// 找到刚输入的段落，设置格式。最后一段
            Dispatch paragraph=Dispatch.call(paragraphs, "Item",new Variant(i+1)).toDispatch();  
            
            int outline=Dispatch.get(paragraph, "OutlineLevel").getInt();  
            Dispatch paraRange1=Dispatch.get(paragraph, "Range").toDispatch();  
            Dispatch listFormat = Dispatch.get(paraRange1, "ListFormat").toDispatch();
            String listString = Dispatch.get(listFormat, "ListString").toString();
            String quanName = Dispatch.get(paraRange1, "Text").toString();
            if (quanName.indexOf("__________________")>=0) {
				continue;
			}else{
				 quan.add(listString+quanName);
		         contentAll.add(quanName);
			} 
            if(outline<=9){  //判断是否为大纲  
            	Dispatch paragraph1=Dispatch.call(paragraphs, "Item",new Variant(i+1)).toDispatch();  
                Dispatch paraRange=Dispatch.get(paragraph1, "Range").toDispatch();
                //根据标签来找到该标签的范围
            	String name = Dispatch.get(paraRange, "Text").toString();
            	//判断当前大纲是空，跳出本次循环
            	if (name.replaceAll("\\\r|\\\f", "").equals("")) {
            		continue;
            	}else{
            		if(name.indexOf("前   言")>=0){
                		dagan.add(name);
                	}else{
                		dagan.add(listString+name);
                	}
//            		.replaceAll("(.)|(\\v)|(.)|(^)","<br>")
//                	title.add(name.replaceAll("(\\v)","<br>"));
                	title.add(name.replaceAll("(\\v)","\n"));
                	if (listString.equals("")){
                		listStringName.add("0");
                	}else{
                		int len = listString.length();
                		String chop = listString.substring(0,len-1);
                    	listStringName.add(chop);
                	}
                	//降级操作
                	
                	if (name.replaceAll("\\\r|\\\f", "").equals("范围")) {
						if (outline==1) {
							llll = 13;
						}else{
							llll = 12;
						}
						
					}
                	if (llll == 12) {
                		if(outline!=1){
	                		outline=outline-1;
	                	}
					}
//                	
                	
            		outlineClass.add(outline);
            	}
            	
            }
         
        }  

       
       int daganLen = 0;
       String daganAll = "";
       String daganTest = "";
       //获取当前术语和定义存在的集合下标
       for (int j = 0; j < dagan.size(); j++) {
    	   if (title.get(j).replaceAll("\\\r|\\\f", "").indexOf("术语与定义")>=0 || title.get(j).replaceAll("\\\r|\\\f", "").indexOf("术语和定义")>=0) {
    		   daganLen = j;
    		   daganAll = dagan.get(j+1);
    		   daganTest = dagan.get(j);
    	   }
    	   
       }
       
       int quanLen = 0;
       String quanAll = "";
       int quanLenEr = 0;
       int jian = 0;
       //获取当前术语和定义在全文的集合下标
       for (int j = 0; j < quan.size(); j++) {
    	   if (quan.get(j).replaceAll("\\\r|\\\f", "").equals(daganTest.replaceAll("\\\r|\\\f", ""))){
    		   quanLen = j;
    		   quanAll = quan.get(j+1);
		   }
    	   
    	   if (quan.get(j).replaceAll("\\\r|\\\f", "").equals(dagan.get(daganLen+1).replaceAll("\\\r|\\\f", ""))){
    		   jian = j-quanLen-1;
    		   quanLenEr = j;
		   }
    	  
       }
       String Text = "123";
       //判断当前术语和定义下面内容的格式：
       if (dagan.get(daganLen+1).indexOf("3.1")<0 && dagan.get(daganLen+2).indexOf("3.2")<0 && dagan.get(daganLen+3).indexOf("3.3")<0) {
    	   if (jian == 1){
    		   oneSplit(quan.get(quanLen+1),quanLen,quanLenEr,daganLen);
    	   }else{
    		   for (int j = 1; j <= 10; j++) {
            	   if(quan.get(quanLen+j).replaceAll("\\\r|\\\f", "").equals("3.1")){
            		   numberMerge(daganLen,daganAll,daganTest);
            		   Text = "1234";
            		   break;
            	   }
               }
               //当前如果是123就代表3.1后面的数据都在一行
               if (Text.equals("123")){
            	   disposeLineFeed(daganLen,daganAll,daganTest);
               }
    	   }
       }
       dislodgeListName(filePath);
       dislodgePicture(filePath);
       quan.add("~~~~~~~~~~~~~~");
       dagan.add("~~~~~~~~~~~~~~");
	    }finally{
        	try{
        		if (wordFile != null) {					
        			Dispatch.call(wordFile, "Close", false);
				}
        		if (word != null){
        			word.invoke("Quit", new Variant[] {});
        		}   
        	} catch (Exception e2) {
//        		log.error("关闭ActiveXComponent异常:"+e2);
        	}
	    }
	    //如果没有这句话，winword.exe进程将不会关闭
	    ComThread.Release();
	    ComThread.quitMainSTA();
	}
	//术语：是多行数据
	public static void numberMerge(int daganLen,String daganAll,String daganTest){
		List<String> Id = new ArrayList<String>();
		//存大纲标题
		List<String> saveAll = new ArrayList<String>();
		//Id+saveAll:拼接
		List<String> outlineAll = new ArrayList<String>();
		List<String> outlineAll11 = new ArrayList<String>();
		//存全文
		List<String> saveFullText = new ArrayList<String>();
		//Id+saveFullText:拼接
		List<String> fullTextId = new ArrayList<String>();
		//等级
		List<Integer> gradeAll = new ArrayList<Integer>();
		 int quanLen = 0;
	       int quanLen1 = 0;
	       for (int j = 0; j < quan.size(); j++) {
	    	   if (quan.get(j).replaceAll("\\\r|\\\f", "").equals(daganTest)){
	    		   quanLen = j;
			   }
	    	   if (quan.get(j).equals(daganAll)){
	    		   quanLen1 = j;
	    		   String ss = "3.";
	    		   int len = 1;
	    		   for (int i = quanLen; i < j; i++) {
	    			   if (quan.get(i).replaceAll("\\\r|\\\f", "").equals("3.1")){
	    				   for (int k = quanLen+1; k < i; k++) {
	    					   fullTextId.add(quan.get(k));
	    					   saveFullText.add(quan.get(k));
	    				   }
	    			   }
	    			   if (quan.get(i).replaceAll("\\\r|\\\f", "").equals(ss+len)){
	    				   Id.add(quan.get(i).replaceAll("\\\r|\\\f", ""));
	    				   ++len;
	    			   }
	    		   }
	    		   Id.add(daganAll.replaceAll("\\\r|\\\f", ""));
			   }
	       }
	       for (int i = 0; i < Id.size()-1; i++) {
	    	   gradeAll.add(2);
	    	   String saveId = Id.get(i); 
	    	   String saveId1 = Id.get(i+1);
	    	   int k = 0;
	    	   int g = 0;
	    	   for (int j = 0; j < quan.size(); j++) {
	    		   if(quan.get(j).replaceAll("\\\r|\\\f", "").equals(saveId)){
	    			   k = j;
	    		   }
	    		   if(quan.get(j).replaceAll("\\\r|\\\f", "").equals(saveId1)){
	    			   for (int j2 = k+1; j2 < j; j2++) {
	    				   saveFullText.add(quan.get(j2));
	    			   }
	    			   fullTextId.add(saveId+"　"+quan.get(k+1));
	    			   for (int j2 = k+2; j2 < j; j2++) {
	    				   fullTextId.add(quan.get(j2));
	    			   }
	    			   saveAll.add(quan.get(k+1));
	    		   }
	    	   }
	       }
	       for (int i = 0; i < Id.size()-1; i++) {
	    	   outlineAll.add(Id.get(i)+"　"+saveAll.get(i));
//	    	   outlineAll11.add(saveAll.get(i));
	       }
	       
	       for (int i = quanLen1-1; i>quanLen; i--) {
		    	  quan.remove(i);
		    	  contentAll.remove(i);
		      }
//	       dagan1.addAll(dagan);
	     //添加全文
		      for (int i = 0; i < fullTextId.size(); i++) {
		    	  quan.add(quanLen+1+i, fullTextId.get(i));
		    	  contentAll.add(quanLen+1+i, saveFullText.get(i));
		      }
		      Id.remove(Id.size()-1);
		      
		    //添加大纲
		      for (int i = 0; i < outlineAll.size(); i++) {
		    	  dagan.add(daganLen+1+i, outlineAll.get(i));
//		    	  dagan1.add(daganLen+1+i, outlineAll11.get(i));
		    	  title.add(daganLen+1+i, saveAll.get(i));
		    	  listStringName.add(daganLen+1+i, Id.get(i));
		    	  outlineClass.add(daganLen+1+i, gradeAll.get(i));
		      }
	}
	
	public static void oneSplit(String yi,int quanLen,int quanLenEr,int daganLen){
		yi = yi + "~~~~~~~";
		String[] s=yi.split("");
		int k = 1;
		String kk = "3.";
		//大纲编号
		List<String> outlineId = new ArrayList<String>();
		//大纲带编号
		List<String> outlineNumber = new ArrayList<String>();
		//大纲不带编号
		List<String> outline = new ArrayList<String>();
		//全文带编号
		List<String> fullTextNumber = new ArrayList<String>();
		//全文不带编号
		List<String> fullText = new ArrayList<String>();
		//等级
		List<Integer> grade = new ArrayList<Integer>();
		//编号
//		List<String> number = new ArrayList<String>();
		for (int i = 0; i < s.length; i++) {
			if (s[i].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", "").equals(kk+k)) {
				outlineId.add(s[i].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""));
				++k;
			}
		}
		int g = 0;
		

		outlineId.add("~~~~~~~");
		for (int ii = 0; ii < outlineId.size()-1; ii++) {
			for (int i = 0; i < s.length; i++) {
				if (outlineId.get(ii).equals(s[i].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", "").replaceAll("\\\r|\\\f", ""))) {
					g = i;
				}
				if (outlineId.get(ii+1).equals(s[i].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", "").replaceAll("\\\r|\\\f", ""))) {
					outlineNumber.add(s[g].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", "")+" "+s[g+1].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""));
					outline.add(s[g+1].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""));
					fullTextNumber.add(s[g].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", "")+" "+s[g+1].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""));
					fullText.add(s[g+1].replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""));
					grade.add(2);	
						for (int j = g; j < i; j++) {
							if (j>=g+2) {
								fullTextNumber.add(s[j]);
								fullText.add(s[j]);
							}
						}
						
				}
			}
		}
		for (int i = quanLenEr-1; i>quanLen; i--) {
	    	  quan.remove(i);
	    	  contentAll.remove(i);
	    }
//		 //添加全文
	      for (int i = 0; i < fullTextNumber.size(); i++) {
	    	  quan.add(quanLen+1+i, fullTextNumber.get(i));
	    	  contentAll.add(quanLen+1+i, fullText.get(i));
	      }
//	      //添加大纲
	      for (int i = 0; i < outlineNumber.size(); i++) {
	    	  dagan.add(daganLen+1+i, outlineNumber.get(i));
	    	  title.add(daganLen+1+i, outline.get(i));
	    	  listStringName.add(daganLen+1+i, outlineId.get(i));
	    	  outlineClass.add(daganLen+1+i, grade.get(i));
	      }
			
	}
	
	
	//术语：一行数据
	public static void disposeLineFeed(int daganLen,String daganAll,String daganTest){
		//临时数据存
	 	//大纲（带编号的）
	 	List<String> outlineId = new ArrayList<String>();
	 	//大纲
		List<String> outlineA = new ArrayList<String>();
		//全文（带编号）
		List<String> fullTextId = new ArrayList<String>();
		//全文
		List<String> fullText = new ArrayList<String>();
		//等级
		List<Integer> grade = new ArrayList<Integer>();
		//编号
		List<String> number = new ArrayList<String>();

	       int quanLen = 0;
	       int quanLen1 = 0;
	       for (int j = 0; j < quan.size(); j++) {
	    	   if (quan.get(j).replaceAll("\\\r|\\\f", "").equals(daganTest.replaceAll("\\\r|\\\f", ""))){
	    		   quanLen = j;
			   }
	    	   if (quan.get(j).equals(daganAll)){
	    		   quanLen1 = j;
	    		   for (int i = quanLen; i < quanLen1; i++) {
	    			   String[] s=quan.get(i).split("");
	    				String[] s1=quan.get(i).split("", 3);
	    				if (s.length<=1){
	    					for (int l = 0; l < s.length; l++) {
	    						fullTextId.add(s[l]);
	    						fullText.add(s[l]);
							}
	    					continue;
	    				}else {
	    					outlineId.add(s[0]+s[1]);
//	    					outlineId.add(s[1]);
	    					outlineA.add(s[1]);
	    					fullTextId.add(s[0]+s[1]);
	    					fullTextId.add(s1[2]);
	    					fullText.add(s[1]);
	    					fullText.add(s1[2]);
	    					number.add(s[0]);
	    					grade.add(2);
						}
	    		   }
			   }
	       }
	       fullTextId.remove(0);
	       fullText.remove(0);
	       String nn = quan.get(quanLen+1);
	       //删除全文的指定内容
	      for (int i = quanLen1-1; i>quanLen; i--) {
	    	  quan.remove(i);
	    	  contentAll.remove(i);
	      }
	      //添加全文
	      for (int i = 0; i < fullTextId.size(); i++) {
	    	  quan.add(quanLen+1+i, fullTextId.get(i));
	    	  contentAll.add(quanLen+1+i, fullText.get(i));
	      }
	      //添加大纲
	      for (int i = 0; i < outlineId.size(); i++) {
	    	  dagan.add(daganLen+1+i, outlineId.get(i));
	    	  title.add(daganLen+1+i, outlineA.get(i));
	    	  listStringName.add(daganLen+1+i, number.get(i));
	    	  outlineClass.add(daganLen+1+i, grade.get(i));
	      }
	}
	
	
	//设置序号
	public static void serial(List<String> gan) {
		serialNumber = new ArrayList<String>();
		int[] yi = new int[20] ;
		for (int i = 0; i < outlineClass.size(); i++) {
			switch (outlineClass.get(i)) {
			case 1:
				yi[0] = yi[0]+1;
				serialNumber.add(yi[0]+"");
				break;
			case 2:
				yi[1] = yi[1]+1;
				serialNumber.add(yi[1]+"");
				break;
			case 3:
				yi[2] = yi[2]+1;
				serialNumber.add(yi[2]+"");
				break;
			case 4:
				yi[3] = yi[3]+1;
				serialNumber.add(yi[3]+"");
				break;
			case 5:
				yi[4] = yi[4]+1;
				serialNumber.add(yi[4]+"");
				break;
			case 6:
				yi[5] = yi[5]+1;
				serialNumber.add(yi[5]+"");
				break;
			case 7:
				yi[6] = yi[6]+1;
				serialNumber.add(yi[6]+"");
				break;
			case 8:
				yi[7] = yi[7]+1;
				serialNumber.add(yi[7]+"");
				break;
			case 9:
				yi[8] = yi[8]+1;
				serialNumber.add(yi[8]+"");
				break;
			default://如果大纲等级超过9。就算自定义大纲等级
				yi[9] = yi[9]+1;
				serialNumber.add(yi[9]+"");
				break;
			}
		}

	}
	
	

	//获取内容.replaceAll("(.)|(\\s)|(.)|(^)","")
		public static void wordT(List<String> gan,List<String> quanwen) {
			nei = new ArrayList<String>();
			nei1 = new ArrayList<String>();
			outlineAll = new ArrayList<String>();
			outlineSerialNumber = new ArrayList<String>();
			countXuhao = new ArrayList<Integer>();
//			String yi = gan.get(0);
//			String er = gan.get(1);
			String yi = null;
			String er = null;
			String n = "";

			int k = 0;
			int g = 0;
			int count = 1;
			for (int w = 0; w < gan.size()-1; w++) {
				
				yi = gan.get(w);
				er = gan.get(w+1);
				k = 0;
				g = 0;
				n = "";
				count = 1;
				for (int i = 0; i < quanwen.size(); i++) {
					if (quanwen.get(i).equals(yi)) {
						k = i;
					}
					g = i;
					if (quanwen.get(k).replaceAll("\\\r|\\\f", "").equals("参考文献")) {
						lenAll = nei1.size()-1;
					}
					if (quanwen.get(i).equals(er)) {
						for (int j = k+1; j < g; j++) {
							if(quanwen.get(j).length()==0){
								continue;
							}else {
								outlineAll.add(title.get(w));
								outlineSerialNumber.add(listStringName.get(w));
								String nn = quanwen.get(j).replaceAll("(.)|(\\v)|(.)|(^)","<br>");
//								String nn1 = nn.replaceAll("(.)|(\\s)|(.)|(^)","");
//								String nn1 = nn.replaceAll("(.)|(\\s)|(\\cL)|(\\cJ)|(\\cM)|(\\cI)|(\\cK)|(.)|(^)","");
								String nn1 = nn.replaceAll("(.)|(\\cL)|(\\cJ)|(\\cM)|(\\cI)|(\\cK)|(.)|(^)","");
//								nei1.add(nn1.replaceAll("(^<br>)|(.<br>)","<br>  "));
								
								nei1.add(nn1);
//								nei1.add(nn.replaceAll("(.)|(\\s)|(\\cL)|(\\cJ)|(\\cM)|(\\cI)|(\\cK)|(.)|(^)",""));
								countXuhao.add(count);
								++count;
							}
						}
						if (yi.replaceAll("\\\r|\\\f", "").equals("前   言") || yi.replaceAll("\\\r|\\\f", "").equals("前言")){
							int lll = outlineAll.size();
							int ll = nei1.size();
							int llll = countXuhao.size();
							int lllll = outlineSerialNumber.size();
							nei1.remove(ll-1);
							outlineAll.remove(lll-1);
							countXuhao.remove(llll-1);
							outlineSerialNumber.remove(lllll-1);
						}
						break;
					}
				
			}
			
			}
			if (lenAll!=0) {
				dislodgeNumber();
			}
			
		}
		
		public static void dislodgeNumber(){

			int lll = lenAll+1;
			int oo =0;
			for (int i = lll; i < nei1.size(); i++) {
				if (nei1.get(i).replaceAll("\\\r|\\\f", "").equals("")) {
					oo = i;
					break;
				}
			}
			for (int i = lll; i < oo; i++) {
				nei1.set(i, nei1.get(i).replaceAll("(\\x5B)", "").replaceAll("(\\x5D)", "").replaceAll("^(\\d)", "").replaceAll("^(\\d)", ""));
			}
			for (int i = oo; i < nei1.size(); i++) {
				nei1.remove(i);
			}
			for (int i = nei1.size()-1; i > 0; i--) {
				if (nei1.get(i).replaceAll("\\\r|\\\f", "").equals("")) {
					nei1.remove(i);
				}else{
					break;
				}
			}
		}
		
	//获取word标题。去除带有FORMTEXT的
	public static String headlieAll(String file){
		String name = "";
		tt(file);
		int l =0;
        String[] s = nn.get(2).split("FORMTEXT");
		for (int i = 0; i < s.length; i++) {
			if (s[i].equals("")) {
				continue;
			}else {
				name = s[i].replace("", "").trim();
			}
		}
		return name;
	}
	
	public static void dislodgeListName(String file){
		tt(file);
		List<String> formName = new ArrayList<String>();
		for (int i = 0; i < nn.size(); i++) {
			if (nn.get(i).substring(0,3).substring(0,1).equals("表")) {
				formName.add(nn.get(i));
			}else if(nn.get(i).substring(0,3).substring(1,2).equals("表")){
				formName.add(nn.get(i));
			}
		}
		int ii = 0;
		for (int j = 0; j < formName.size(); j++) {
			for (int i = 0; i < contentAll.size(); i++) {
				if (contentAll.get(i).replaceAll("\\\r|\\\f", "").equals(formName.get(j).replaceAll("\\\r|\\\f", ""))) {
					contentAll.remove(i);
					quan.remove(i);
					break;
				}
				
			}
		}
		String n3 = ".*[A-Za-z][.][0-9].*";
		String n4 = ".*[0-9][.][0-9].*";
		String n5 = ".*[表][0-9].*";
		String n6 = ".*[表][A-Za-z].*";
		for (int i = contentAll.size()-1; i >=0 ; i--) {
			if (contentAll.get(i).length()>6) {
				String nnn = contentAll.get(i).substring(0,5);
				//正则表达式字符串判断
				Matcher m3=Pattern.compile(n3).matcher(nnn);
				Matcher m4=Pattern.compile(n4).matcher(nnn);
				Matcher m5=Pattern.compile(n5).matcher(nnn);
				Matcher m6=Pattern.compile(n6).matcher(nnn);
				if (nnn.indexOf("表")>=0) {
					if (nnn.indexOf(".")>=0) {
						if (m3.matches()) {
							contentAll.remove(i);
							quan.remove(i);
						}else if(m4.matches()){
							contentAll.remove(i);
							quan.remove(i);
						}
					}else{
						if (m5.matches()) {
							contentAll.remove(i);
							quan.remove(i);
						}else if(m6.matches()){
							contentAll.remove(i);
							quan.remove(i);
						}
					}
				}
			}else{
				continue;
			}
		}
		
	}
	//获取word标题
	 public static void tt(String file) {
	    	try {
	          nn = new ArrayList<String>();
	          InputStream is = new FileInputStream(new File(file));  //需要将文件路更改为word文档所在路径。
	          POIFSFileSystem fs = new POIFSFileSystem(is);
	          HWPFDocument document = new HWPFDocument(fs);
	          Range range = document.getRange();

	          CharacterRun run1 = null;//用来存储第一行内容的属性
	          CharacterRun run2 = null;//用来存储第二行内容的属性
	          int q=1;
	          for (int i = 0; i < range.numParagraphs()-1; i++) {
	              Paragraph para1 = range.getParagraph(i);// 获取第i段
	              Paragraph para2 = range.getParagraph(i+1);// 获取第i段
	              int t=i;              //记录当前分析的段落数
	              
	              String paratext1 = para1.text().trim().replaceAll("\r\n", "");   //当前段落和下一段
	              String paratext2 = para2.text().trim().replaceAll("\r\n", "");
	              run1=para1.getCharacterRun(0);
	              run2=para2.getCharacterRun(0);
	              if (paratext1.length() > 0&&paratext2.length() > 0) {
	                      //这个if语句为的是去除大标题，连续三个段落字体大小递减就跳过
	                      if(run1.getFontSize()>run2.getFontSize()&&run2.getFontSize()>range.getParagraph(i+2).getCharacterRun(0).getFontSize()) {
	                          continue;
	                      }                        
	                      //连续两段字体格式不同
	                      if(run1.getFontSize()>run2.getFontSize()) {
	                          
	                          String content=paratext2;
	                          run1=run2;  //从新定位run1  run2
	                          run2=range.getParagraph(t+2).getCharacterRun(0);
	                          t=t+1;
	                          while(run1.getFontSize()==run2.getFontSize()) {
	                              //连续的相同
	                              content+=range.getParagraph(t-1).text().trim().replaceAll("\r\n", "");
	                              run1=run2;
	                              run2=range.getParagraph(t-1).getCharacterRun(0);
	                              t++;
	                          }
	                          
	                          if(paratext1.indexOf("HYPERLINK")==-1&&content.indexOf("HYPERLINK")==-1) {
	                          	nn.add(paratext1);
	                              i=t;
	                              q++;
	                          }
	                              
	                      }
	              }
	          }
	         

	          
	      } catch (Exception e) {
//	          log.error("异常："+e);
	      }
	 }
	
	//查询大纲id
		public static List<String> T004getID(List<String> name,String domainname,List<String> outlineSerialNumber,String connector){
			List<String> id = new ArrayList<String>();
			id = T004subscriptIdSelect(domainname, name,outlineSerialNumber,connector);
		     return id;
		}
		//插入大纲
		public static void addOutline(String file,String connector,String domainname){
			long lenA = 0;
			long plus = 0;
//			//调用Outlin类：获取大纲相关数据
//			outline.moldAll(file);
			//获取word标题
			String name = headlieAll(file);
			//获取上级id
			String kkk = "";
			String jjj = "";
			String[] superior = new String[11];
			for (int i = 0; i < dagan.size(); i++) {
				
				//获取自增id
				String number1 = getBatchId(domainname,"eiis.seq_eiis_rulesmanage_t003");
				
				if (title.get(i).indexOf("前")>=0 && title.get(i).indexOf("言")>=0){
					//添加大纲
					addOutline(domainname,number1,"0",connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "0", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
					superior[0] = number1;
				}else if(listStringName.get(i).indexOf("附")>=0 && listStringName.get(i).indexOf("录")>=0){
					addOutline(domainname,number1,"0",connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "2", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
					superior[0] = number1;
					kkk = "~~~";
					jjj = superior[0];
				}else if(title.get(i).indexOf("文献")>=0){
					addOutline(domainname,number1,"0",connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "3", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
					superior[0] = number1;
				}else if(title.get(i).indexOf("术语")>=0 && title.get(i).indexOf("定义")>=0){
					addOutline(domainname,number1,"0",connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "4", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
					superior[0] = number1;
					superior[10] = number1;
				}else if(listStringName.get(i).indexOf("3.")>=0 &&(listStringName.get(i).length()==3 || listStringName.get(i).length()==4)){
					addOutline(domainname,number1,superior[10],connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "4", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
				}else if(kkk.equals("~~~")&&outlineClass.get(i)==2){
					addOutline(domainname,number1,jjj,connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "2", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
					superior[1] = number1;
				}else{
					switch (outlineClass.get(i)) {
					case 1:
						addOutline(domainname,number1,"0",connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[0] = number1;
						break;
					case 2:
						addOutline(domainname,number1,superior[0],connector,serialNumber.get(i), title.get(i), outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[1] = number1;
						break;
					case 3:
						addOutline(domainname,number1,superior[1],connector, serialNumber.get(i), title.get(i),outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[2] = number1;
						break;
					case 4:
						addOutline(domainname,number1,superior[2],connector,serialNumber.get(i), title.get(i),outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[3] = number1;
						break;
					case 5:
						addOutline(domainname,number1,superior[3],connector,serialNumber.get(i),title.get(i),outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[4] = number1;
						break;
					case 6:
						addOutline(domainname,number1,superior[4],connector,serialNumber.get(i),title.get(i),outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[5] = number1;
						break;
					case 7:
						addOutline(domainname,number1,superior[5],connector,serialNumber.get(i),title.get(i),outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[6] = number1;
						break;
					case 8:
						addOutline(domainname,number1,superior[6],connector, serialNumber.get(i), title.get(i), outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[7] = number1;
						break;
					case 9:
						addOutline(domainname,number1,superior[7],connector,serialNumber.get(i),title.get(i),outlineClass.get(i), "1", "1", "1", "1", "1", "1", "1", "1", "1",listStringName.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""),BATCHID);
						superior[8] = number1;
						break;
					}
				}
				
			}
		}
		
		//插入内容
		public static void addT004Content(String domainname,String connector,String filepath) throws IOException{	
			
			dagan.add("~~~~~~~~~~~~~~");
			//查询当前内容对应的大纲id和序号
			List<String> xuhaoId = T004getID(outlineAll,domainname,outlineSerialNumber,connector);
			String number="";
			int uu = 0;
			for (int i = 0; i < nei1.size(); i++) {
				//判断如果字符超出2000就换成二段内容插入
				if (nei1.get(i).length()>=2000){
					String cutOut = nei1.get(i).substring(0,1999);
					String cutOutRear = nei1.get(i).substring(1999,nei1.get(i).length());
					number = getBatchId(domainname,"eiis.seq_eiis_rulesmanage_t004");
					addContentT004(domainname,number, xuhaoId.get(i),countXuhao.get(i)+"",cutOut, "0","1","1","1","1","1","1","1","1",BATCHID);
					number = getBatchId(domainname,"eiis.seq_eiis_rulesmanage_t004");
					addContentT004(domainname,number, xuhaoId.get(i), countXuhao.get(i)+"",cutOutRear, "0","1","1","1","1","1","1","1","1",BATCHID);
				}else{
					number = getBatchId(domainname,"eiis.seq_eiis_rulesmanage_t004");
					addContentT004(domainname,number, xuhaoId.get(i), countXuhao.get(i)+"",nei1.get(i), "0","1","1","1","1","1","1","1","1",BATCHID);
				}
			}
		}

		/*
		 * GUICHENGXMBM:项目编号（id）
		 * GUICHENGBM:参数（前端传过来的接口）
		 * SHANGJIXMBM：上级项目编号（之前的id）
		 * XUHAO：序号
		 * XIANGMUMC:规程项目名称
		 * MULUCJ:目录层级
		 * LEIXING:类型（类型 0:前言 1:标题 2:附录 3:文献 4:术语与定义）
		 * */
	    public static  void addOutline(String domainName,String GUICHENGXMBM,String SHANGJIXMBM,String GUICHENGBM,String XUHAO,String XIANGMUMC,int MULUCJ,String LEIXING,String GONGXING,String GONGSI,String GEZHOUB,String SANXIA,String XIANGJIAB,String XILUOD,String BAIHET,String WUDONGD,String OUTLINELEVEL,String BATCHID){
	    	 
	    	Connection connection = null;
	    	PreparedStatement pstmt = null;
	 	 
	 	   	try {
	 	   		connection = RulesManageUtils.getConnection(domainName);
				String sql = "insert into eiis.EIIS_RULESMANAGE_T003(GUICHENGXMBM,SHANGJIXMBM,XUHAO,XIANGMUMC,MULUCJ,LEIXING,GONGXING,GONGSI,GEZHOUB,SANXIA,XIANGJIAB,XILUOD,BAIHET,WUDONGD,GUICHENGBM,BATCHID,STATUS,OUTLINELEVEL) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
				pstmt = connection.prepareStatement(sql);
				pstmt.setString(1, GUICHENGXMBM);
				pstmt.setString(2, SHANGJIXMBM);
				pstmt.setString(3, XUHAO);
				pstmt.setString(4, XIANGMUMC.replace("", "").trim());
				pstmt.setInt(5, MULUCJ);
				pstmt.setString(6, LEIXING);
				pstmt.setString(7, GONGXING);
				pstmt.setString(8, GONGSI);
				pstmt.setString(9, GEZHOUB);
				pstmt.setString(10, SANXIA);
				pstmt.setString(11, XIANGJIAB);
				pstmt.setString(12, XILUOD);
				pstmt.setString(13, BAIHET);
				pstmt.setString(14, WUDONGD);
				pstmt.setString(15, GUICHENGBM);
				pstmt.setString(16, BATCHID);
				pstmt.setString(17, "0");
				pstmt.setString(18, OUTLINELEVEL);
				pstmt.executeUpdate();//运行增删改操作
			} catch (Exception e) {
				// TODO Auto-generated catch block
//				log.error("异常："+e);
			}finally {
		        if (pstmt != null) {
		            try {
		            	pstmt.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
		        if (connection != null) {
		            try {
		                connection.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
			}
		}
	    
	    public static String getBatchId(String domainName,String sequenceName){
	    	Connection connection = null;
	    	PreparedStatement pstmt = null;
	    	ResultSet rs = null;
	  		String len = "";
	  		try {
	  			connection = RulesManageUtils.getConnection(domainName);
	  			String sql = "SELECT " +sequenceName+".nextval as C from dual";
	  			pstmt = connection.prepareStatement(sql);
	  			rs = pstmt.executeQuery();
	  			while(rs.next()){
	  				len = rs.getString("c");
	  			}
	  			
	  		} catch (Exception e) {
	  			// TODO Auto-generated catch block
//	  			log.error("异常："+e);
	  		}finally{
	  			if (rs != null) {
	  	            try {
	  	                rs.close();
	  	            } catch (SQLException e) {
//	  	                log.error("异常："+e);
	  	            }
	  	        }
	  	        if (pstmt != null) {
	  	            try {
	  	            	pstmt.close();
	  	            } catch (SQLException e) {
//	  	                log.error("异常："+e);
	  	            }
	  	        }
	  	        if (connection != null) {
	  	            try {
	  	                connection.close();
	  	            } catch (SQLException e) {
//	  	                log.error("异常："+e);
	  	            }
	  	        }
	  			}  		
	  		return len;
	  	}
	    
	   
	    
		public static  void addContentT004(String domainName,String NEIRONGBH,String GUICHENGXMBM,String XUHAO,String NEIRONG,String NEIRONGLX,String GONGXING,String GONGSI,String GEZHOUB,String SANXIA,String XIANGJIAB,String XILUOD,String BAIHET,String WUDONGD,String BATCHID){
			Connection connection = null;
	    	PreparedStatement pstmt = null;
	    	ResultSet rs = null;
			try {
				connection = RulesManageUtils.getConnection(domainName);
				String sql = "insert into eiis.EIIS_RULESMANAGE_T004(NEIRONGBH,guichengxmbm,xuhao,neirong,neironglx,gongxing,gongsi,gezhoub,sanxia,xiangjiab,xiluod,baihet,wudongd,BATCHID,STATUS) values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
				pstmt = connection.prepareStatement(sql);
				pstmt.setString(1, NEIRONGBH);
				pstmt.setString(2, GUICHENGXMBM);
				pstmt.setString(3, XUHAO);
				pstmt.setString(4, NEIRONG);
				pstmt.setString(5, NEIRONGLX);
				pstmt.setString(6, GONGXING);
				pstmt.setString(7, GONGSI);
				pstmt.setString(8, GEZHOUB);
				pstmt.setString(9, SANXIA);
				pstmt.setString(10, XIANGJIAB);
				pstmt.setString(11, XILUOD);
				pstmt.setString(12, BAIHET);
				pstmt.setString(13, WUDONGD);
				pstmt.setString(14, BATCHID);
				pstmt.setString(15, "0");
				pstmt.executeUpdate();//运行增删改操作
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
//				log.error("异常："+e);
			}finally{
				if (rs != null) {
		            try {
		                rs.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
		        if (pstmt != null) {
		            try {
		            	pstmt.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
		        if (connection != null) {
		            try {
		                connection.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
			}
		}
		
				

		public static List<String> T004subscriptIdSelect(String domainName,List<String> name,List<String> outlineSerialNumber,String GUICHENGBM){
	  		List<String> subscript = new ArrayList<String>();
	  		String text = "";
	  		String xuhao = "";
	  		Connection connection = null;
	    	PreparedStatement pstmt = null;
	    	ResultSet rs = null;
//	    	outlineSerialNumber.remove(0);
	  		try {
	  			connection = RulesManageUtils.getConnection(domainName);
	  			String sql = "select GUICHENGXMBM as id from eiis.EIIS_RULESMANAGE_T003 where GUICHENGXMBM = (select MAX(cast(GUICHENGXMBM as int)) as id from eiis.EIIS_RULESMANAGE_T003 where xiangmumc=? and outlinelevel =? and GUICHENGBM=?)";
	  			pstmt = connection.prepareStatement(sql);
	  			for (int i = 0; i < name.size(); i++) {
	  				pstmt.setString(1, name.get(i).replace("", "").trim());
	  				pstmt.setString(2, outlineSerialNumber.get(i).replaceAll("^[　 ]*", "").replaceAll("[　 ]*$", ""));
	  				pstmt.setString(3, GUICHENGBM);
	  				rs = pstmt.executeQuery();
	  	  			while(rs.next()){
	  	  				subscript.add(rs.getString("ID"));
	  	  			}
				}
	  				
	  			
	  		} catch (Exception e) {
	  			// TODO Auto-generated catch block
//	  			log.error("异常："+e);
	  		}finally{
	  			if (rs != null) {
	  	            try {
	  	                rs.close();
	  	            } catch (SQLException e) {
//	  	                log.error("异常："+e);
	  	            }
	  	        }
	  	        if (pstmt != null) {
	  	            try {
	  	            	pstmt.close();
	  	            } catch (SQLException e) {
//	  	                log.error("异常："+e);
	  	            }
	  	        }
	  	        if (connection != null) {
	  	            try {
	  	                connection.close();
	  	            } catch (SQLException e) {
//	  	                log.error("异常："+e);
	  	            }
	  	        }		
	  			}  		
	  		return subscript;
	  	}
	  	
	  	public static  void updateBatchid(String domainName,String batchid){
			Connection connection = null;
	    	PreparedStatement pstmt = null;
	    	ResultSet rs = null;
			try {
				connection = RulesManageUtils.getConnection(domainName);
				String sql = "update eiis.EIIS_RULESMANAGE_T003 set batchid = '',status='1' where batchid=?";
				pstmt = connection.prepareStatement(sql);
				pstmt.setString(1, batchid);
			
				pstmt.executeUpdate();//运行增删改操作
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
//				log.error("异常："+e);
			}finally{
				if (rs != null) {
		            try {
		                rs.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
		        if (pstmt != null) {
		            try {
		            	pstmt.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
		        if (connection != null) {
		            try {
		                connection.close();
		            } catch (SQLException e) {
//		                log.error("异常："+e);
		            }
		        }
			}
		}
}
