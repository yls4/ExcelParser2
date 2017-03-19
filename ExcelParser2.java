import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;


//Import the JExcel API
import jxl.Workbook;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import jxl.Cell;

// Only works on Excel 97-2003 files.  If you receive an error, save the excel file as a 97-2003 file and try again.
public class ExcelParser2 {
	static Workbook workbook;
	static WritableWorkbook copy;
	
	public static void main(String[] args) throws BiffException, IOException, WriteException {
		Scanner reader = new Scanner(System.in);
		System.out.println("Enter file name");
		String fileName = reader.nextLine();
		try {
			workbook = Workbook.getWorkbook(new File(fileName));
			copy = Workbook.createWorkbook(new File("temp.xls"), workbook);
			WritableSheet copySheet = copy.getSheet(0);
			WritableWorkbook output = Workbook.createWorkbook(new File("output.xls"));
			WritableSheet wsheet = output.createSheet("Sheet1", 0);
			//FileReader input = new FileReader(wordList);
			//BufferedReader list = new BufferedReader(input);
			//String myLine = null;
			//ArrayList<String> words = new ArrayList<String>();
			
			//while ((myLine = list.readLine()) != null) {
				//words.add(myLine);
			//}
			
			//WorkSheet sheet = workbook.getSheet(1);
			WritableSheet sheet2 = copy.getSheet(0);
			int rows = sheet2.getRows();
			int cols = sheet2.getColumns();
			Label c1 = new Label(0, 0, "Mall Name");
			Label c2 = new Label(1, 0, "Address");
			Label c3 = new Label(2, 0, "Town");
			Label c4 = new Label(3, 0, "State");
			Label c5 = new Label(4, 0, "Zipcode");
			Label c6 = new Label(5, 0, "GLA");
			Label c7 = new Label(6, 0, "Description");
			Label c8 = new Label(7, 0, "Contact First");
			Label c9 = new Label(8, 0, "Contact Last");
			Label c10 = new Label(9, 0, "Email");
			Label c11 = new Label(10, 0, "Phone");
			wsheet.addCell(c1);
			wsheet.addCell(c2);
			wsheet.addCell(c3);
			wsheet.addCell(c4);
			wsheet.addCell(c5);
			wsheet.addCell(c6);
			wsheet.addCell(c7);
			wsheet.addCell(c8);
			wsheet.addCell(c9);
			wsheet.addCell(c10);
			wsheet.addCell(c11);

			System.out.println(copySheet.getCell(2, 0).getContents());
			for (int row=1; row<rows; row++) {
				Label col1 = new Label(0, row, copySheet.getCell(0, row).getContents());
				String[] s = copySheet.getCell(2, row).getContents().split("\\r?\\n");
				Label col2 = new Label(1, row, s[2]);
				String[] s2 = s[3].split(",");
				Label col3 = new Label(2, row, s2[0]);
				String ss = s2[1].replaceAll("[^A-Za-z\\s]", "");
				Label col4 = new Label(3, row, ss.trim());
				String zip = s2[1].replaceAll("[^\\d]", "");
				Label col5 = new Label(4, row, zip.trim());
				String gla = "";
				for (String temp: s) {
					temp = temp.replaceAll("[^\\w\\s]:\\(\\)", "");
					if (temp.contains("GLA")) {
						gla = temp.split(":")[1];
						//System.out.println(gla);
					}
				}
				Label col6 = new Label(5, row, gla.trim());
				String desc = copySheet.getCell(3, row).getContents();
				Label col7 = new Label(6, row, desc);
				String[] s4 = copySheet.getCell(4, row).getContents().split("\\r?\\n");
				s4[0] = s4[0].replaceAll("[^\\w\\s]", "");
				String[] sss = s4[0].split("\\s+");
				String first = sss[0];
				String last = "";
				if (sss.length > 1) {
					last = sss[sss.length-1];
				}
				Label col8 = new Label(7, row, first);
				Label col9 = new Label(8, row, last);
				
				String email = "";
				for (String temp: s4) {
					if (temp.contains("@")) {
						email = temp;
					}
				}
				
				Label col10 = new Label(9, row, email);
				
				String phone = s4[s4.length-1];
				Label col11 = new Label(10, row, phone);
				
				wsheet.addCell(col1);
				wsheet.addCell(col2);
				wsheet.addCell(col3);
				wsheet.addCell(col4);
				wsheet.addCell(col5);
				wsheet.addCell(col6);
				wsheet.addCell(col7);
				wsheet.addCell(col8);
				wsheet.addCell(col9);	
				wsheet.addCell(col10);
				wsheet.addCell(col11);
			}
			
			copy.write();
			copy.close();
			output.write();
			output.close();
			System.out.println("Output stored in temp.xls");
		}
		catch(Error E) {
			System.out.println("File not found");
		}
	}
}
