package poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *	解析一个指定xls的数据
 * @author Administrator
 *
 */
public class PoiExcelTest {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//		
		//		try (InputStream inp = new FileInputStream("students.xlsx")) {
		//			//InputStream inp = new FileInputStream("workbook.xlsx");
		//			    Workbook wb = WorkbookFactory.create(inp);
		//			    Sheet sheet = wb.getSheetAt(0);
		//			    Row row = sheet.getRow(2);
		//			    Cell cell = row.getCell(3);
		//			    if (cell == null)
		//			        cell = row.createCell(3);
		//			    cell.setCellType(CellType.STRING);
		//			    cell.setCellValue("a test");
		//			    // Write the output to a file
		//			    try (OutputStreamream fileOut = new FileOutputStream("workbook.xls")) {
		//			        wb.write(fileOut);
		//			    }
		//			}
		//读取数据流
		InputStream inp = new FileInputStream("aa.xls");

		//解析工作簿
		HSSFWorkbook workBook=new HSSFWorkbook(inp);

		//解析工作表得到sheet的数量
		int size=workBook.getNumberOfSheets();
		System.out.println("表中一共有"+size+"个sheet");

		//创建一个sheet
		HSSFSheet sheet;
		//创建一个row
		HSSFRow row;
		//创建一个cell
		HSSFCell cell;

		//循环处理每一个工作表中的数据
		for (int i = 0; i < size; i++) {
			//拿到第i个sheet
			sheet=workBook.getSheetAt(i);
			System.out.println("工作表sheet的名字："+sheet.getSheetName());

			//通过每个sheet得到它有效行数,有数据的行数
			int rowNumber=sheet.getPhysicalNumberOfRows();

			System.out.println("共有"+rowNumber+"行数据");

			//循环操作每行数据
			for (int rowIndex = 0; rowIndex < rowNumber; rowIndex++) {
				System.out.println("正在读取第"+rowIndex+"行数据");

				//拿到第rowIndex行
				row=sheet.getRow(rowIndex);

				//拿到每行的列数
				int cellNumber=row.getPhysicalNumberOfCells();
				System.out.println("第"+rowIndex+"行有"+cellNumber+"列数据");


				for (int cellIndex = 0; cellIndex < cellNumber; cellIndex++) {

					//拿到第rowIndex行的第cellIndex列
					cell=row.getCell(cellIndex);

					switch (cell.getCellType()) {
					case HSSFCell.CELL_TYPE_STRING:
						System.out.println("-----------"+cell.getStringCellValue());
						break;
					case HSSFCell.CELL_TYPE_NUMERIC:
						System.out.println("-----------"+cell.getNumericCellValue());
					default:
						break;
					}
				}
			}
		}
	}
}
