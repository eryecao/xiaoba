package com.xiaoba.lsj;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * ��Windows Server 2008�ϣ��Ҳ���dll�ļ�������Exception in thread "main" java.lang.UnsatisfiedLinkError: no jacob-1.17-x86 in java.library.path
 * ��win7��win8�ϳɹ�ʵ�֡�
 * 
 * @author Administrator
 *
 */
public class TestDoc {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String path="D:\\yanqiong.doc";
		System.out.println("��ʼ��ӡ");
		ComThread.InitSTA();
		ActiveXComponent word=new ActiveXComponent("Word.Application");
		Dispatch doc=null;
		Dispatch.put(word, "Visible", new Variant(false));
		Dispatch docs=word.getProperty("Documents").toDispatch();
		doc=Dispatch.call(docs, "Open", path).toDispatch();
		
		try {
			Dispatch.call(doc, "PrintOut");//��ӡ
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("��ӡʧ��");
		}finally{
			try {
				if(doc!=null){
					Dispatch.call(doc, "Close",new Variant(0));
				}
			} catch (Exception e2) {
				e2.printStackTrace();
			}
			//�ͷ���Դ
			ComThread.Release();
		}
		
		
		
	
	}

}
