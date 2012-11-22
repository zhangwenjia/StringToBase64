import java.util.ArrayList;

public class Main {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		ArrayList<String> emails = new ArrayList<String>();
		ArrayList<String> base64emails = new ArrayList<String>();
		
		ConvertBase64Util util64 = new ConvertBase64Util();
		emails = util64.readFromExcel();
		base64emails = util64.convertToBase64(emails);
		util64.writeToExcel(base64emails);
	}

}
