package cn.miaozw.showExcel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;

public class ShowTask implements Runnable {
    private final JTextArea jTextArea;
    private final File file;

    public ShowTask(JTextArea jTextArea, File file) {
        this.jTextArea = jTextArea;
        this.file = file;
    }

    @Override
    public void run() {
        try {
            showExcel(file,jTextArea);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
    }

    private static void showExcel(File file, JTextArea msgTextArea) throws IOException, InvalidFormatException, InterruptedException {
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 行对象
        for (int ri = sheet.getFirstRowNum() + 1; ri <= sheet.getLastRowNum(); ri++) {
            msgTextArea.setText("");
            msgTextArea.append("\t月份\t\t");
            msgTextArea.append("\t数据\t\t\n\n");
            Row row = sheet.getRow(ri);
            for (int ci = row.getFirstCellNum(); ci < row.getLastCellNum(); ci++) {
                Cell cell = row.getCell(ci);
                Object value = getCellValueByType(cell);
                System.out.printf("%s ", value);
                msgTextArea.append("\t" + value.toString() + "\t\t");
            }
            Thread.sleep(1000);
        }
        workbook.close();
    }

    private static Object getCellValueByType(Cell cell) {
        String value;
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                //如果为时间格式的内容
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    //注：format格式 yyyy-MM-dd hh:mm:ss 中小时为12小时制，若要24小时制，则把小h变为H即可，yyyy-MM-dd HH:mm:ss
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    value = sdf.format(HSSFDateUtil.getJavaDate(cell.
                            getNumericCellValue()));
                    break;
                } else {
                    value = new DecimalFormat("0").format(cell.getNumericCellValue());
                }
                break;
            case HSSFCell.CELL_TYPE_STRING: // 字符串
                value = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                value = cell.getBooleanCellValue() + "";
                break;
            case HSSFCell.CELL_TYPE_FORMULA: // 公式
                value = cell.getCellFormula() + "";
                break;
            case HSSFCell.CELL_TYPE_BLANK: // 空值
                value = "";
                break;
            case HSSFCell.CELL_TYPE_ERROR: // 故障
                value = "非法字符";
                break;
            default:
                value = "未知类型";
                break;
        }
        return value;
    }
}
