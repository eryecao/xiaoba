package com.xiaoba.lsj;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * 在Windows Server 2008上，找不到dll文件，报错：Exception in thread "main" java.lang.UnsatisfiedLinkError: no jacob-1.17-x86 in java.library.path
 * 在win7，win8上成功实现。
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
		System.out.println("开始打印");
		ComThread.InitSTA();
		ActiveXComponent word=new ActiveXComponent("Word.Application");
		Dispatch doc=null;
		Dispatch.put(word, "Visible", new Variant(false));
		Dispatch docs=word.getProperty("Documents").toDispatch();
		doc=Dispatch.call(docs, "Open", path).toDispatch();
		
		try {
			Dispatch.call(doc, "PrintOut");//打印
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("打印失败");
		}finally{
			try {
				if(doc!=null){
					Dispatch.call(doc, "Close",new Variant(0));
				}
			} catch (Exception e2) {
				e2.printStackTrace();
			}
			//释放资源
			ComThread.Release();
		}
		
		
		
	
	}

}
