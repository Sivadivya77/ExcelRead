package excelread;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readexcel1 {

	public static void main(String[] args) throws Exception {

		Readexcel1 x = new Readexcel1();

		for (int i = 0; i < 5; i++) {
			for (int j = 0; j < 3; j++)
				System.out.println(x.getExecelData("sheet1", i, j) + " ");
		}

	}

	public String getExecelData(String Sheetname, int rownum, int colunum) {
		String retval = null;

		try {
			FileInputStream fis = new FileInputStream("until//Students.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet s = wb.getSheet(Sheetname);
			XSSFRow r = s.getRow(rownum);
			XSSFCell c = r.getCell(colunum);
			retval = c.getStringCellValue();
			fis.close();
			wb.close();

		} catch (FileNotFoundException e) {

			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return retval;

	}

	public static String getcellvalue(XSSFCell c) {
		switch (c.getCellType()) {
		case NUMERIC:
			return String.valueOf(c.getNumericCellValue());
		case BOOLEAN:
			return String.valueOf(c.getBooleanCellValue());
		case STRING:
			return c.getStringCellValue();
		default:
			return c.getStringCellValue();

		}

	}
}
/*output 
 * Name age Email
 * john Dos	30	john@test.com
Jane Dos	28	Jane@test.com
Bob Smith	35	Bob@test.com
Swapnil	37	Swapnil@test.com*/

