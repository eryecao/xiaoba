package com.xiaoba.lsj;

import com.jacob.activeX.*;
import com.jacob.com.*;

public class Print {
	//private static final String PRINT_NAME=new String("FX7AF20D"); 
	/**
	 * ˵��
	 *  <p>��1�����Ƚ�JACOB��DLL������ C:\Windows\System32 ��</p>
	 *  <p>��2��DCom Server Process Launcher��DcomLaunch��������Ҫ��</p>
	 *  <p>��3����Ҫ��װMicrosoft office 2003+</p>
	 *  <p>��ӡ���Բμ� <a href="http://msdn.microsoft.com/zh-cn/library/office/ff838253.aspx" target="_blank">PrintOut ���� (Excel)</a>
	 * @param path ��ӡ·����ַ������ \\XX\\YY.xls
	 * @param copies ��ӡ����
	 */
	public static void printExcel(String path,int copies){
		if(path.isEmpty()||copies<1){
			return;
		}
		//��ʼ��COM�߳�
		ComThread.InitSTA();
		//�½�Excel����
		ActiveXComponent xl=new ActiveXComponent("Excel.Application");
		try { 
			System.out.println("Version=" + xl.getProperty("Version"));
			//�����Ƿ���ʾ��Excel  
			Dispatch.put(xl, "Visible", new Variant(true));
			//�򿪾���Ĺ�����
			Dispatch workbooks = xl.getProperty("Workbooks").toDispatch(); 
			Dispatch excel=Dispatch.call(workbooks,"Open",System.getProperty("user.dir")+path).toDispatch(); 
			
			//���ô�ӡ���Բ���ӡ
			Dispatch.callN(excel,"PrintOut",new Object[]{Variant.VT_MISSING, Variant.VT_MISSING, new Integer(copies),
					new Boolean(false),/*PRINT_NAME*/Variant.VT_MISSING, new Boolean(true),Variant.VT_MISSING, ""});
			
			//�ر��ĵ�
			//Dispatch.call(excel, "Close", new Variant(false));  
		} catch (Exception e) { 
			e.printStackTrace(); 
		} finally{
			//xl.invoke("Quit",new Variant[0]);
			//ʼ���ͷ���Դ 
			ComThread.Release(); 
		} 
	}
}
