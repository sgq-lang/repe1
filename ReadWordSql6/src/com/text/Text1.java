package com.text;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class Text1 {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		kk("E:\\����\\�ڲ���\\11\\����\\22.doc");
	}
	
	public static boolean xmlToWord(String docfile, String htmlfile) {
		boolean flag = false;
		ComThread.InitSTA();
		int WORD_DOC = 0;
//		int WORD_DOCX = 12;
		ActiveXComponent app = null;
		try {
			app = new ActiveXComponent("Word.Application"); // ����word
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
			/* ����Ҫ���ҵ����� */
			Dispatch.put(find, "Text", "Ŀ   ��");
			/* ��ǰ���� */
			Dispatch.put(find, "Forward", "True");
			/* ���ø�ʽ */
			Dispatch.put(find, "Format", "True");
			/* ��Сдƥ�� */
			Dispatch.put(find, "MatchCase", "True");
			/* ȫ��ƥ�� */
			Dispatch.put(find, "MatchWholeWord", "True");
			/* ���Ҳ�ѡ�� */
			Dispatch.call(find, "Execute").getBoolean();
			/* ȡ��ActiveDocument��TablesOfContents��range���� */
			Dispatch ActiveDocument = app.getProperty("ActiveDocument")
					.toDispatch();
			Dispatch TablesOfContents = Dispatch.get(ActiveDocument,
					"TablesOfContents").toDispatch();
			Dispatch.call(selection, "MoveRight"); // �ƶ���굽�ұ�
			Dispatch.call(selection, "TypeParagraph"); // ����
			Dispatch range = Dispatch.get(selection, "Range").toDispatch();

			/****************************/

			Dispatch pageSetup = Dispatch.get(doc,"PageSetup").toDispatch();
			Dispatch.put(pageSetup, "OddAndEvenPagesHeaderFooter", new Variant(true));
			
			//ȡ�û�������
			Dispatch activeWindow = app.getProperty( "ActiveWindow").toDispatch();
			//ȡ�û�������
		    Dispatch activePane = Dispatch.get(activeWindow, "ActivePane").toDispatch();			
		    //ȡ���Ӵ�����
		    Dispatch view = Dispatch.get(activePane, "View").toDispatch();
		    
		    // ��ҳü��ֵΪ9��ҳ��Ϊ10
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
		    
		    Dispatch headerFooterRange = Dispatch.get(headerFooter,"Range").toDispatch(); //��ǰѡ�е�ҳü����
		    
		    Dispatch paragraphs =Dispatch.get(headerFooterRange,"Paragraphs").getDispatch();
		    //Dispatch.put(paragraphs , "Alignment", new Variant(3));	 
		    //Dispatch.put(headerFooterRange,"Text","2");
		    
		    String content = Dispatch.get(headerFooterRange,"Text").toString(); //��õ�ǰҳü�е�����
		    //replace(docSelection,"");

		    //System.out.println(content.equals("\t"));
		    
		    System.out.println("i==,content==="+content+"jjj");			
		    
		    //�ر�ҳü
	        Dispatch.put(view, "SeekView", new Variant(0));		    
	
			/****************************/		
			/* ����Ŀ¼ */
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
			// �رս���
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
		//ȡ�û�������
		Dispatch oSelection = oWord.getProperty("Selection").toDispatch();
		Dispatch selection = oWord.getProperty("Selection").toDispatch();//��ö�Selection���
		Dispatch.call(selection, "HomeKey", new Variant(6));//�Ƶ���ͷ
		Dispatch find = Dispatch.call(selection, "Find").toDispatch();//���Find���
		Dispatch.put(find, "Text", "�����ɹ����"); //�����ַ���"�����ɹ����"
		Dispatch.call(find, "Execute"); //ִ�в�ѯ

		 String pages = Dispatch.call(selection, "Information",new Variant(3)).toString();

		System.out.println("�ı�����ҳ��:"+pages);
		        Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
				/**���Ϊ*/
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
		//ȡ�û�������
		Dispatch selection = oWord.getProperty("Selection").toDispatch();
		Dispatch ActiveDocument = oWord.getProperty("ActiveDocument")
				.toDispatch();
		Dispatch TablesOfContents = Dispatch.get(ActiveDocument,
				"TablesOfContents").toDispatch();
		Dispatch.call(selection, "MoveRight"); // �ƶ���굽�ұ�
		Dispatch.call(selection, "TypeParagraph"); // ����
		Dispatch range = Dispatch.get(selection, "Range").toDispatch();
				//ȡ�û�������
				Dispatch activeWindow = oWord.getProperty("ActiveWindow").toDispatch();
				Dispatch activePan = Dispatch.get(activeWindow, "ActivePane").toDispatch();
				 // ȡ���Ӵ�����
		        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
		     // ��ҳü��ֵΪ9��ҳ��Ϊ10
		        Dispatch.put(view, "SeekView", new Variant(10));
//		        Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
//			    Dispatch.call(selection,"MoveDown");
			    
		        Dispatch docSelection = Dispatch.get(activeWindow, "Selection").toDispatch();
			    
			    Dispatch headerFooter = Dispatch.get(docSelection, "HeaderFooter").toDispatch();
			    
			    Dispatch headerFooterRange = Dispatch.get(headerFooter,"Range").toDispatch(); //��ǰѡ�е�ҳü����
			    Dispatch fields = Dispatch.get(range, "Fields").toDispatch();
			    Dispatch paragraphFormat=Dispatch.get(selection,"ParagraphFormat").getDispatch();
			    Dispatch.put(paragraphFormat, "Alignment", 0);
			  //Dispatch.call(fields, "Add", new Variant(range), new Variant(-1), new Variant(""), new Variant("True")) .toDispatch();
			  Dispatch.call(fields, "Add", Dispatch.get(selection, "Range").toDispatch(), new Variant(-1), "Page", true).toDispatch();

//			  Dispatch.call(fields, "Add", Dispatch.get(selection, "Range").toDispatch(), new Variant(-1), "NumPages",true).toDispatch();
//			  Dispatch font = Dispatch.get(range, "Font").toDispatch();
//			  Dispatch.put(font,"Name",new Variant("����_GB2312"));
//			  Dispatch.put(font, "Bold", new Variant(true));
//			  Dispatch.put(font, "Size", 9);
		        //�ر�ҳü
		        Dispatch.put(view, "SeekView", new Variant(0));
		        Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
				/**���Ϊ*/
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
		//ȡ�û�������
		Dispatch selection = oWord.getProperty("Selection").toDispatch();
		Dispatch ActiveDocument = oWord.getProperty("ActiveDocument")
				.toDispatch();
		Dispatch TablesOfContents = Dispatch.get(ActiveDocument,
				"TablesOfContents").toDispatch();
		Dispatch.call(selection, "MoveRight"); // �ƶ���굽�ұ�
		Dispatch.call(selection, "TypeParagraph"); // ����
		Dispatch range = Dispatch.get(selection, "Range").toDispatch();
		Dispatch pageSetup = Dispatch.get(oDocument,"PageSetup").toDispatch();
		Dispatch.put(pageSetup, "OddAndEvenPagesHeaderFooter", new Variant(true));
				//ȡ�û�������
				Dispatch activeWindow = oWord.getProperty("ActiveWindow").toDispatch();
				Dispatch activePan = Dispatch.get(activeWindow, "ActivePane").toDispatch();
				 // ȡ���Ӵ�����
		        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
		     // ��ҳü��ֵΪ9��ҳ��Ϊ10
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
			    	Dispatch headerFooterRange = Dispatch.get(headerFooter,"Range").toDispatch(); //��ǰѡ�е�ҳü����
//				    Dispatch paragraphs =Dispatch.get(headerFooterRange,"Paragraphs").getDispatch();
				    String content = Dispatch.get(headerFooterRange,"Text").toString(); //��õ�ǰҳü�е�����
				    if (!content.replaceAll("\\\r|\\\f", "").equals("")) {
				    	Dispatch.call(docSelection,"MoveDown");
				        Dispatch.call(docSelection,"MoveDown");
//				        Dispatch.call(docSelection,"MoveDown");
//				        Dispatch.call(docSelection,"MoveDown");
//				    	System.out.println("�հ�");
					}else{
						break;
					}
				}
			    Dispatch paragraphs = Dispatch.get(selection, "Paragraphs").toDispatch();
			      Dispatch.put(paragraphs, "Alignment", new Variant(0)); // ���뷽ʽ
			      Dispatch fields = Dispatch.get(range, "Fields").toDispatch();
//			      Dispatch.call(fields, "Add", new Variant(range), new Variant(-1), new Variant(""), new Variant("True")) .toDispatch();
			      Dispatch.call(fields, "Add", Dispatch.get(selection, "Range").toDispatch(), new Variant(-1), "Page", true).toDispatch();
			    
		        //�ر�ҳü
		        Dispatch.put(view, "SeekView", new Variant(0));
		        Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
				/**���Ϊ*/
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
		
		//ȡ�û�������

		Dispatch activeWindow = oWord.getProperty("ActiveWindow").toDispatch();
		
		//ȡ�û�������

		Dispatch activePan = Dispatch.get(activeWindow, "ActivePane").toDispatch();
		 // ȡ���Ӵ�����
        Dispatch view = Dispatch.get(activePan, "View").toDispatch();
     // ��ҳü��ֵΪ9��ҳ��Ϊ10
        Dispatch.put(view, "SeekView", new Variant(10));
        Dispatch docSelection = Dispatch.get(activeWindow, "Selection").toDispatch();
        //��ȡҳü��ҳ��
        Dispatch headfooter = Dispatch.get(docSelection, "HeaderFooter").toDispatch();
        // ��ȡˮӡͼ�ζ���
        Dispatch shapes = Dispatch.get(headfooter, "Shapes").toDispatch();
        // ���ĵ�ȫ������ˮӡ,������ˮӡЧ�������ݣ����壬��С���Ƿ�Ӵ֣��Ƿ�б�壬��߾࣬�ϱ߾ࡣ
//        Dispatch paragraphs = Dispatch.get(docSelection, "Paragraphs").toDispatch();
//        Dispatch.put(paragraphs, "Alignment", new Variant(0)); // ���뷽ʽ
        Dispatch.call(docSelection, "MoveLeft");
      //�ر�ҳü
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

		//ȡ�û���ݶ���

		Dispatch ActiveDocument = oWord.getProperty("ActiveDocument").toDispatch();

		//ȡ�û�������

		Dispatch ActiveWindow = oWord.getProperty("ActiveWindow").toDispatch();

		//ȡ�û�������

		Dispatch ActivePane = Dispatch.get(ActiveWindow, "ActivePane").toDispatch();

		//ȡ���Ӵ�����

		Dispatch View = Dispatch.get(ActivePane, "View").toDispatch();

		Dispatch.put(View, "SeekView", "10");
      Dispatch paragraphs = Dispatch.get(oSelection, "Paragraphs").toDispatch();
      Dispatch.put(paragraphs, "Alignment", new Variant(0)); // ���뷽ʽ
		Dispatch.put(oSelection, "Text", "ҳü������ɣ��� ");

//		Dispatch.call(oSelection, "MoveLeft");
		
		Dispatch.call((Dispatch) Dispatch.call(oWord, "WordBasic").getDispatch(),"FileSaveAs", file);
		/**���Ϊ*/
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
	    //���û����仰��winword.exe���̽�����ر�
	}

}
