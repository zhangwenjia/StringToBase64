import it.sauronsoftware.base64.Base64;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ConvertBase64Util {

	private ArrayList<String> emails = new ArrayList<String>();
	private ArrayList<String> base64emails = new ArrayList<String>();

	private static final String PATH = "D:/StringToBase64/";
	private static final String FILE_NAME = "users.xls";
	private static final String FILE_NAME1= "users1.xls";
	
	/**
	 * 从excel中读取数据
	 * 
	 * @return
	 */
	public ArrayList<String> readFromExcel() {
		Workbook workbook;
		try {
			workbook = Workbook.getWorkbook(new File(PATH + FILE_NAME));
			Sheet sheet = workbook.getSheet(0);
			int rows = sheet.getRows();

			for (int i = 0; i < rows; i++) {
				for (int j = 0; j < 1; j++) {
					Cell cellTemp = sheet.getCell(j, i);
					String email = cellTemp.getContents().trim();
					emails.add(email);
				}
			}

			workbook.close();
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return emails;
	}

	/**
	 * 把原始字符串转成base64编码格式的字符串
	 * 
	 * @param emails
	 * @return
	 */
	public ArrayList<String> convertToBase64(ArrayList<String> emails) {
		int size = emails.size();
		for (int i = 0; i < size; i++) {
			String email = emails.get(i);
			String base64email = Base64.encode(email);
			base64emails.add(base64email);
		}

		for (int k = 0; k < size; k++) {
			System.out.println("base64email is:" + base64emails.get(k));
		}

		return base64emails;
	}

	/**
	 * 将编码后的字符串写入excel
	 * 
	 * @param base64emails
	 * @throws BiffException
	 */
	public void writeToExcel(ArrayList<String> base64emails) {
		try {

			WritableWorkbook workbook = Workbook.createWorkbook(new File(PATH + FILE_NAME1));
			WritableSheet sheet = workbook.createSheet("sheet1", 0);

			int size = base64emails.size();
			for (int i = 0; i < size; i++) {
				String base64email = base64emails.get(i);

				Label l = new Label(1, i, base64email);
				sheet.addCell(l);
			}

			workbook.write();
			workbook.close();

		} catch (IOException e) {
			e.printStackTrace();
		} catch (RowsExceededException e) {
			e.printStackTrace();
		} catch (WriteException e) {
			e.printStackTrace();
		}
	}

}
