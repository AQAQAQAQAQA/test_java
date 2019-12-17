
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;

public class bb {
    public static void main(String[] args) throws Exception {
        File xlsFile = new File("C:\\Users\\31675\\Desktop\\2019年出纳日记帐【昊居】（大良招行--、番禺交行）-测试(1)(1).xlsx");

        // 工作表
        Workbook workbook = WorkbookFactory.create(xlsFile);

        // 表个数。
        int numberOfSheets = workbook.getNumberOfSheets();

        // 遍历表。
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);

            // 行数。
            int rowNumbers = sheet.getLastRowNum() + 1;

            // Excel第一行。
            Row temp = sheet.getRow(0);
            if (temp == null) {
                continue;
            }

            int cells = temp.getPhysicalNumberOfCells();

            // 读数据。
            for (int row = 0; row < rowNumbers; row++) {
                Row r = sheet.getRow(row);
                for (int col = 0; col < cells; col++) {
                    System.out.print(r.getCell(col).toString() + " ");
                }

                // 换行。
                System.out.println();
            }
        }
    }
}
