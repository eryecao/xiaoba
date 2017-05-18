package com.xiaoba.lsj;

import java.io.IOException;
import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * Servlet implementation class PrintServlet
 */
@WebServlet("/PrintServlet")
public class PrintServlet extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public PrintServlet() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.getWriter().append("Served at: ").append(request.getContextPath());
		
		String path="G:\\59store店长管理系统_操作手册.docx";
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

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}

}
