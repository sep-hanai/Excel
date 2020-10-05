package servlet;

import java.io.File;
import java.io.IOException;

import javax.servlet.ServletException;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Servlet implementation class TopBL
 */
@WebServlet("/TopBL")
public class TopBL extends HttpServlet {
	private static final long serialVersionUID = 1L;

    /**
     * @see HttpServlet#HttpServlet()
     */
    public TopBL() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
    //エクセルをダウンロードするサーブレット
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		response.getWriter().append("Served at: ").append(request.getContextPath());
		 //エクセルファイルへアクセスするためのオブジェクト
	    Workbook excel = WorkbookFactory.create(new File("Sample.xlsx"));

	    // シート名がわかっている場合
	    Sheet sheet = excel.getSheet("Sheet1");

	    //0行目
	    Row row = sheet.getRow(0);

	    //0番目のセル
	    Cell cell = row.getCell(0);

	    //文字列の取得
	    String value = cell.getStringCellValue();

	    //取得した文字列の表示
	    System.out.println(value);
		//ダウンロードページTop.jspを表示する
		getServletContext().getRequestDispatcher("/Top.jsp").forward(request, response);
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
//		String filename = "Sample.xlsx";//Excelファイル名
//		//レスポンスヘッダを設定
//		response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");//ファイルのタイプを.xlsxに指定
//        response.setHeader("content-disposition", String.format("attachment; filename=\"%s\"", filename));//ファイル名指定
//        response.setCharacterEncoding("UTF-8");//ファイルの文字コード変換

		// TODO Auto-generated method stub
		doGet(request, response);
	}

}
