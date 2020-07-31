package com.text;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Text1 {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		kk("E:\\测试\\在测试\\11\\测试\\22.doc");
	}
	
	public static boolean xmlToWord(String docfile, String htmlfile) {
		boolean flag = false;
		ComThread.InitSTA();
		int WORD_DOC = 0;
//		int WORD_DOCX = 12;
		ActiveXComponent app = null;
		try {
			app = new ActiveXComponent("Word.Application"); // 启动word
			app.setProperty("Visible", new Variant(false));
			Dispatch docs = app.getProperty("Documents").toDispatch();
			Dispatch doc = Dispatch.invoke(
					docs,
					"Open",
					Dispatch.Method,
					new Object[] { docfile, new Variant(false),
							new Variant(true) }, new int[1]).toDispatch();
			Dispatch selection = app.getProperty("Selection").toDispatch();
			Dispatch find = Dispatch.call(selection, "Find").toDispatch();
			/* 设置要查找的内容 */
			Dispatch.put(find, "Text", "目   次");
			/* 向前查找 */
			Dispatch.put(find, "Forward", "True");
			/* 设置格式 */
			Dispatch.put(find, "Format", "True");
			/* 大小写匹配 */
			Dispatch.put(find, "MatchCase", "True");
			/* 全字匹配 */
			Dispatch.put(find, "MatchWholeWord", "True");
			/* 查找并选中 */
			Dispatch.call(find, "Execute").getBoolean();
			/* 取得ActiveDocument、TablesOfContents、range对象 */
			Dispatch ActiveDocument = app.getProperty("ActiveDocument")
					.toDispatch();
			Dispatch TablesOfContents = Dispatch.get(ActiveDocument,
					"TablesOfContents").toDispatch();
			Dispatch.call(selection, "MoveRight"); // 移动光标到右边
			Dispatch.call(selection, "TypeParagraph"); // 换行
			Dispatch range = Dispatch.get(selection, "Range").toDispatch();

			/****************************/

			Dispatch pageSetup = Dispatch.get(doc,"PageSetup").toDispatch();
			Dispatch.put(pageSetup, "OddAndEvenPagesHeaderFooter", new Variant(true));
			
			//取得活动窗体对象
			Dispatch activeWindow = app.getProperty( "ActiveWindow").toDispatch();
			//取得活动窗格对象
		    Dispatch activePane = Dispatch.get(activeWindow, "ActivePane").toDispatch();			
		    //取得视窗对象
		    Dispatch view = Dispatch.get(activePane, "View").toDispatch();
		    
		    // 打开页眉，值为9，页脚为10
		    Dispatch.put(view, "SeekView", "10");
		    
		    Dispatch.call(selection,"MoveDown" );
		    Dispatch.call(selection,"MoveDown" );
		    Dispatch.call(selection,"MoveDown" );
		    Dispatch.call(selection,"MoveDown" );
		    Dispatch.call(selection,"MoveDown" );
		    Dispatch.call(selection,"MoveDown" );
//		    Dispatch.call(selection,"MoveDown" );
//		    Dispatch.call(selection,"MoveDown" );
		   
		    
		    Dispatch docSelection = Dispatch.get(activeWindow, "Selection").toDispatch();
		    
		    Dispatch headerFooter = Dispatch.get(docSelection, "HeaderFooter").toDispatch();
		    
		    Dispatch headerFooterRange = Dispatch.get(headerFooter,"Range").toDispatch(); //当前选中的页眉对象
		    
		    Dispatch paragraphs =Dispatch.get(headerFooterRange,"Paragraphs").getDispatch();
		    //Dispatch.put(paragraphs , "Alignment", new Variant(3));	 
		    //Dispatch.put(headerFooterRange,"Text","2");
		    
		    String content = Dispatch.get(headerFooterRange,"Text").toString(); //获得当前页眉中的内容
		    //replace(docSelection,"");

		    //System.out.println(content.equals("\t"));
		    
		    System.out.println("i==,content==="+content+"jjj");			
		    
		    //关闭页眉
	        Dispatch.put(view, "SeekView", new Variant(0));		    
	
			/****************************/		
			/* 增加目录 */
			Dispatch add = Dispatch.call(TablesOfContents, "Add", range,
					new Variant(true), new Variant(1), new Variant(2),
					new Variant(true), new Variant(true), new Variant('T'),
					new Variant(true), new Variant(true), new Variant(1),
					new Variant(false)).toDispatch();
			Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] {
					htmlfile, new Variant(WORD_DOC) }, new int[1]);

			Variant f = new Variant(false);

			Dispatch.call(doc, "Close", f);

			flag = true;
		} catch (Exception e) {
			flag = false;
//			log.error(e.getMessage());
		} finally {
			if (app != null)
				app.invoke("Quit", new Variant[] {});
			// 关闭进程
			ComThread.Release();
		}
		return flag;
	}
	public static void kk1(String file){
		ActiveXComponent oWord = null;
		Dispatch oDocument = null;
		try{
		oWord = new ActiveXComponent("Word.Application");

		oWord.setProperty("Visible", new Variant(false));

		Dispatch oDocuments = oWord.getProperty("Documents").toDispatch();

		oDocument = Dispatch.invoke(oDocuments, "Open", Dispatch.Method, new Object[]{file,new Variant(true),new Variant(false)}, new int[1]).toDispatch();
		//取得活动窗体对象
		Dispatch oSelection = oWord.getProperty("Selection").toDispatch();
		Dispatch selection = oWord.getProperty("Selection").toDispatch();//获得对Selection组件
		Dispatch.call(selection, "HomeKey", new Variant(6));//移到开头
		Dispatch find = Dispatch.call(selection, "Find").toDispatch();//获得Find组件
		Dispatch.put(find, "Text", "二、成果简介"); //查找字符串"二、成果简介"
		Dispatch.call(find, "Execute"); //执行查询

		 String pages = Dispatch.call(selection, "Information",new Variant(3)).toString();

		System.out.println("文本所在页码:"+pages);
		        Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
				/**另存为*/
				Dispatch.invoke(oDocument, "SaveAs", Dispatch.Method, new Object[] {    
						file, new Variant(true) }, new int[1]);
		}finally{
        	try{
        		if (oDocument != null) {					
        			Dispatch.call(oDocument, "Close", false);
				}
        		if (oWord != null){
        			oWord.invoke("Quit", new Variant[] {});
        		}   
        	} catch (Exception e2) {
        		
        	}
	    }
	}
	
	public static void kk2(String file){
		ActiveXComponent oWord = null;
		Dispatch oDocument = null;
		try{
		oWord = new ActiveXComponent("Word.Application");

		oWord.setProperty("Visible", new Variant(false));

		Dispatch oDocuments = oWord.getProperty("Documents").toDispatch();

		oDocument = Dispatch.invoke(oDocuments, "Open", Dispatch.Method, new Object[]{file,new Variant(true),new Variant(false)}, new int[1]).toDispatch();
		//取得活动窗体对象
		Dispatch selection = oWord.getProperty("Selection").toDispatch();
		Dispatch ActiveDocument = oWord.getProperty("ActiveDocument")
				.toDispatch();
		Dispatch TablesOfContents = Dispatch.get(ActiveDocument,
				"TablesOfContents").toDispatch();
		Dispatch.call(selection, "MoveRight"); // 移动光标到右边
		Dispatch.call(selection, "TypeParagraph"); // 换行
		Dispatch range = Dispatch.get(selection, "Range").toDispatch();
				//取得活动窗格对象
				Dispatch activeWindow = oWord.getProperty("ActiveWindow").toDispatch();
				Dispatch activePan = Dispatch.get(activeWindow, "ActivePane").toDispatch();
				 // 取得视窗对象
		        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
		     // 打开页眉，值为9，页脚为10
		        Dispatch.put(view, "SeekView", new Variant(10));
//		        Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
			    
		        Dispatch docSelection = Dispatch.get(activeWindow, "Selection").toDispatch();
			    
			    Dispatch headerFooter = Dispatch.get(docSelection, "HeaderFooter").toDispatch();
			    
			    Dispatch headerFooterRange = Dispatch.get(headerFooter,"Range").toDispatch(); //当前选中的页眉对象
			    Dispatch fields = Dispatch.get(range, "Fields").toDispatch();
			    Dispatch paragraphFormat=Dispatch.get(selection,"ParagraphFormat").getDispatch();
			    Dispatch.put(paragraphFormat, "Alignment", 0);
			  //Dispatch.call(fields, "Add", new Variant(range), new Variant(-1), new Variant(""), new Variant("True")) .toDispatch();
			  Dispatch.call(fields, "Add", Dispatch.get(selection, "Range").toDispatch(), new Variant(-1), "Page", true).toDispatch();

//			  Dispatch.call(fields, "Add", Dispatch.get(selection, "Range").toDispatch(), new Variant(-1), "NumPages",true).toDispatch();
//			  Dispatch font = Dispatch.get(range, "Font").toDispatch();
//			  Dispatch.put(font,"Name",new Variant("楷体_GB2312"));
//			  Dispatch.put(font, "Bold", new Variant(true));
//			  Dispatch.put(font, "Size", 9);
		        //关闭页眉
		        Dispatch.put(view, "SeekView", new Variant(0));
		        Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
				/**另存为*/
				Dispatch.invoke(oDocument, "SaveAs", Dispatch.Method, new Object[] {    
						file, new Variant(true) }, new int[1]);
		}finally{
        	try{
        		if (oDocument != null) {					
        			Dispatch.call(oDocument, "Close", false);
				}
        		if (oWord != null){
        			oWord.invoke("Quit", new Variant[] {});
        		}   
        	} catch (Exception e2) {
        		
        	}
	    }
	}
	
	
	public static void kk(String file){
		ActiveXComponent oWord = null;
		Dispatch oDocument = null;
		try{
		oWord = new ActiveXComponent("Word.Application");

		oWord.setProperty("Visible", new Variant(false));

		Dispatch oDocuments = oWord.getProperty("Documents").toDispatch();

		oDocument = Dispatch.invoke(oDocuments, "Open", Dispatch.Method, new Object[]{file,new Variant(true),new Variant(false)}, new int[1]).toDispatch();
		//取得活动窗体对象
		Dispatch selection = oWord.getProperty("Selection").toDispatch();
		Dispatch ActiveDocument = oWord.getProperty("ActiveDocument")
				.toDispatch();
		Dispatch TablesOfContents = Dispatch.get(ActiveDocument,
				"TablesOfContents").toDispatch();
		Dispatch.call(selection, "MoveRight"); // 移动光标到右边
		Dispatch.call(selection, "TypeParagraph"); // 换行
		Dispatch range = Dispatch.get(selection, "Range").toDispatch();
		Dispatch pageSetup = Dispatch.get(oDocument,"PageSetup").toDispatch();
		Dispatch.put(pageSetup, "OddAndEvenPagesHeaderFooter", new Variant(true));
				//取得活动窗格对象
				Dispatch activeWindow = oWord.getProperty("ActiveWindow").toDispatch();
				Dispatch activePan = Dispatch.get(activeWindow, "ActivePane").toDispatch();
				 // 取得视窗对象
		        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
		     // 打开页眉，值为9，页脚为10
		        Dispatch.put(view, "SeekView", new Variant(10));
		        
//		        Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
			    
		        Dispatch docSelection = Dispatch.get(activeWindow, "Selection").toDispatch();
		        Dispatch.call(docSelection,"MoveDown");
		        Dispatch.call(docSelection,"MoveDown");
		        Dispatch.call(docSelection,"MoveDown");
		        Dispatch.call(docSelection,"MoveDown");
		        Dispatch.call(docSelection,"MoveDown");
		        Dispatch.call(docSelection,"MoveDown");
//		        Dispatch.call(docSelection,"MoveDown");
//		        Dispatch.call(docSelection,"MoveDown");
			    Dispatch headerFooter = Dispatch.get(docSelection, "HeaderFooter").toDispatch();
			    
			    
			    for (int i = 0; i < 100; i++) {
			    	Dispatch headerFooterRange = Dispatch.get(headerFooter,"Range").toDispatch(); //当前选中的页眉对象
//				    Dispatch paragraphs =Dispatch.get(headerFooterRange,"Paragraphs").getDispatch();
				    String content = Dispatch.get(headerFooterRange,"Text").toString(); //获得当前页眉中的内容
				    if (!content.replaceAll("\\\r|\\\f", "").equals("")) {
				    	Dispatch.call(docSelection,"MoveDown");
				        Dispatch.call(docSelection,"MoveDown");
//				        Dispatch.call(docSelection,"MoveDown");
//				        Dispatch.call(docSelection,"MoveDown");
//				    	System.out.println("空白");
					}else{
						break;
					}
				}
			    Dispatch paragraphs = Dispatch.get(selection, "Paragraphs").toDispatch();
			      Dispatch.put(paragraphs, "Alignment", new Variant(0)); // 对齐方式
			      Dispatch fields = Dispatch.get(range, "Fields").toDispatch();
//			      Dispatch.call(fields, "Add", new Variant(range), new Variant(-1), new Variant(""), new Variant("True")) .toDispatch();
			      Dispatch.call(fields, "Add", Dispatch.get(selection, "Range").toDispatch(), new Variant(-1), "Page", true).toDispatch();
			    
		        //关闭页眉
		        Dispatch.put(view, "SeekView", new Variant(0));
		        Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
				/**另存为*/
				Dispatch.invoke(oDocument, "SaveAs", Dispatch.Method, new Object[] {    
						file, new Variant(true) }, new int[1]);
		}finally{
        	try{
        		if (oDocument != null) {					
        			Dispatch.call(oDocument, "Close", false);
				}
        		if (oWord != null){
        			oWord.invoke("Quit", new Variant[] {});
        		}   
        	} catch (Exception e2) {
        		
        	}
	    }
	}
	
	public static void nn(String file){
		ActiveXComponent oWord = null;
		Dispatch oDocument = null;
		try{
		oWord = new ActiveXComponent("Word.Application");

		oWord.setProperty("Visible", new Variant(true));

		Dispatch oDocuments = oWord.getProperty("Documents").toDispatch();

		oDocument = Dispatch.invoke(oDocuments, "Open", Dispatch.Method, new Object[]{file,new Variant(true),new Variant(false)}, new int[1]).toDispatch();
		
		//取得活动窗体对象

		Dispatch activeWindow = oWord.getProperty("ActiveWindow").toDispatch();
		
		//取得活动窗格对象

		Dispatch activePan = Dispatch.get(activeWindow, "ActivePane").toDispatch();
		 // 取得视窗对象
        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
     // 打开页眉，值为9，页脚为10
        Dispatch.put(view, "SeekView", new Variant(10));
        Dispatch docSelection = Dispatch.get(activeWindow, "Selection").toDispatch();
        //获取页眉和页脚
        Dispatch headfooter = Dispatch.get(docSelection, "HeaderFooter").toDispatch();
        // 获取水印图形对象
        Dispatch shapes = Dispatch.get(headfooter, "Shapes").toDispatch();
        // 给文档全部加上水印,设置了水印效果，内容，字体，大小，是否加粗，是否斜体，左边距，上边距。
//        Dispatch paragraphs = Dispatch.get(docSelection, "Paragraphs").toDispatch();
//        Dispatch.put(paragraphs, "Alignment", new Variant(0)); // 对齐方式
        Dispatch.call(docSelection, "MoveLeft");
      //关闭页眉
        Dispatch.put(view, "SeekView", new Variant(0));
		
		}finally{
        	try{
        		if (oDocument != null) {					
        			Dispatch.call(oDocument, "Close", false);
				}
        		if (oWord != null){
        			oWord.invoke("Quit", new Variant[] {});
        		}   
        	} catch (Exception e2) {
        		
        	}
	    }
		
	}
	
	public static void mm(String file){
		ActiveXComponent oWord = null;
		Dispatch oDocument = null;
		try{
		oWord = new ActiveXComponent("Word.Application");

		oWord.setProperty("Visible", new Variant(true));

		Dispatch oDocuments = oWord.getProperty("Documents").toDispatch();

		oDocument = Dispatch.invoke(oDocuments, "Open", Dispatch.Method, new Object[]{file,new Variant(true),new Variant(false)}, new int[1]).toDispatch();

		Dispatch oSelection = oWord.getProperty("Selection").toDispatch();

		Dispatch oFind = oWord.call(oSelection, "Find").toDispatch();

		Dispatch alignment = Dispatch.get(oSelection, "ParagraphFormat").toDispatch();

		Dispatch image = Dispatch.get(oSelection, "InLineShapes").toDispatch();

		//取得活动内容对象

		Dispatch ActiveDocument = oWord.getProperty("ActiveDocument").toDispatch();

		//取得活动窗体对象

		Dispatch ActiveWindow = oWord.getProperty("ActiveWindow").toDispatch();

		//取得活动窗格对象

		Dispatch ActivePane = Dispatch.get(ActiveWindow, "ActivePane").toDispatch();

		//取得视窗对象

		Dispatch View = Dispatch.get(ActivePane, "View").toDispatch();

		Dispatch.put(View, "SeekView", "10");
      Dispatch paragraphs = Dispatch.get(oSelection, "Paragraphs").toDispatch();
      Dispatch.put(paragraphs, "Alignment", new Variant(0)); // 对齐方式
		Dispatch.put(oSelection, "Text", "页眉你出来吧！！ ");

//		Dispatch.call(oSelection, "MoveLeft");
		
		Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
		/**另存为*/
		Dispatch.invoke(oDocument, "SaveAs", Dispatch.Method, new Object[] {    
				file, new Variant(true) }, new int[1]);
		}finally{
        	try{
        		if (oDocument != null) {					
        			Dispatch.call(oDocument, "Close", false);
				}
        		if (oWord != null){
        			oWord.invoke("Quit", new Variant[] {});
        		}   
        	} catch (Exception e2) {
        		
        	}
	    }
	    //如果没有这句话，winword.exe进程将不会关闭
	}

}
