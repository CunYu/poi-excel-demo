import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

public class UserModelSSReadDemo {

    public static void main(String[] args) throws IOException {

        // 获得Excel操作对象
        try (Workbook wb = new XSSFWorkbook("F:\\Demo.xlsx")) {

            // 遍历Sheet
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {

                // 获得当前Sheet
                Sheet sheet = wb.getSheetAt(i);
                // 获得当前Sheet的名字
                String sheetName = sheet.getSheetName();
                System.out.println(sheetName);

                // 遍历当前Sheet的行
                for (int j = 0; j <= sheet.getLastRowNum(); j++) {

                    // 获得当前行
                    Row row = sheet.getRow(j);

                    // 遍历当前行的单元格
                    for (int k = 0; k < row.getLastCellNum(); k++) {

                        // 获得当前单元格
                        Cell cell = row.getCell(k);

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
}