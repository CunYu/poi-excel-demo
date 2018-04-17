import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class UserModelXSSFReadDemo {

    public static void main(String[] args) throws IOException {

        // 获得Excel操作对象
        XSSFWorkbook xwb = new XSSFWorkbook("F:\\Demo.xlsx");

        // 遍历Sheet
        for (int i = 0; i < xwb.getNumberOfSheets(); i++) {

            // 获得当前Sheet
            XSSFSheet sheet = xwb.getSheetAt(i);
            // 获得当前Sheet的名字
            String sheetName = sheet.getSheetName();
            System.out.println(sheetName);

            // 遍历当前Sheet的行
            for (int j = 0; j <= sheet.getLastRowNum(); j++) {

                // 获得当前行
                XSSFRow row = sheet.getRow(j);

                // 遍历当前行的单元格
                for (int k = 0; k < row.getLastCellNum(); k++) {

                    // 获得当前单元格
                    XSSFCell cell = row.getCell(k);

                    // 当前单元格为字符串格式
                    if (CellType.STRING == cell.getCellTypeEnum()) {
                        String stringContent = cell.getStringCellValue();
                        System.out.print(stringContent + " ");
                    }

                    // 当前单元格为数字格式
                    if (CellType.NUMERIC == cell.getCellTypeEnum()) {
                        double doubleContent = cell.getNumericCellValue();
                        System.out.print((int) doubleContent + " ");
                    }
                }
                System.out.println();
            }
        }
    }
}