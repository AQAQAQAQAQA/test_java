import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * @Author: HLT
 * @Date: Created in 9:16 2019/10/14
 * @Description: excel(出纳日记帐, 读取A3,A4,A5作为新的三列, 作为excel新加入的三列)
 */
public class excel_util {

    public static List<Map<String, String>> getList(String filePath ) {
        Workbook wb = null;
        Sheet sheet = null;
        Row row = null;
        List<Map<String, String>> list = null;
        String cellData = null;
//        String filePath = "C:\\Users\\31675\\Desktop\\2019年出纳日记帐【昊居】（大良招行--、番禺交行）-测试(1)(1).xlsx";
        String columns[] = { "kv" };
        wb = readExcel(filePath);
        if (wb != null) {
            // 用来存放表中数据
            list = new ArrayList<Map<String, String>>();
            // 获取第一个sheet
            sheet = wb.getSheetAt(0);
            // 获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            // 获取第3,4,5行
//            row = sheet.getRow(0);
            // 获取最大列数
//            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 2; i < 5; i++) {
                Map<String, String> map = new HashMap<String, String>();
                row = sheet.getRow(i);
                if (row != null) {
                    for (int j = 0; j < 1; j++) {
                        cellData = (String) getCellFormatValue(row.getCell(0));
                        map.put(columns[0], cellData);
                    }
                } else {
                    break;
                }
                list.add(map);
            }
        }
        return list;
    }


    // 读取excel
    public static Workbook readExcel(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if (".xls".equals(extString)) {
                return wb = new HSSFWorkbook(is);
            } else if (".xlsx".equals(extString)) {
                return wb = new XSSFWorkbook(is);
            } else {
                return wb = null;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return wb;
    }

    public static Object getCellFormatValue(Cell cell) {
        Object cellValue = null;
        if (cell != null) {
            // 判断cell类型
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC: {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA: {
                    // 判断cell是否为日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    } else {
                        // 数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }



    public static void main(String[] args) {
        String filePath = "C:\\Users\\31675\\Desktop\\2019年出纳日记帐【昊居】（大良招行--、番禺交行）-测试(1)(1).xlsx";

        List<Map<String, String>> list= getList(filePath);

        Map<String,String> hashMap = new HashMap<String, String>();
        for (Map<String, String> map : list) {
            if (map.values().toString().contains("：")) {
                hashMap.put(map.values().toString().split("：")[0].replace("[", ""),
                        map.values().toString().split("：")[1].replace("]", ""));
            }else if(map.values().toString().contains(":"))  {
                hashMap.put(map.values().toString().split(":")[0].replace("[", ""),
                        map.values().toString().split(":")[1].replace("]", ""));
            }
        }
        String zhannghu = hashMap.get("账户");
        String zhannghao = hashMap.get("账号");
        String gongsi = hashMap.get("公司");
        System.out.println(zhannghu);
        System.out.println(zhannghao);
        System.out.println(gongsi);
//        Iterator iter = hashMap.entrySet().iterator();
//        while (iter.hasNext()) {
//            Map.Entry entry = (Map.Entry) iter.next();
//            String key = entry.getKey().toString();
//            String value = entry.getValue().toString();
//            System.out.println(key);
//            System.out.println(value);
//
//        }

    }
}
