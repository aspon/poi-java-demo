package cn.huangmp.poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by huangMP on 2017/8/20.
 * decription :
 */
public class ExcelStyledDemo {



    /**
     * 合并单元格 : 合并单元格对象是属于工作簿(独立创建);运用于工作表
     * 居中 new CellRangeAddress(2, 2, 2, 4 ); 构造参数 起始行号 结束行号 起始列号 结束列号
     * 样式属于工作簿, 运用于单元格;
     *  字体属于工作簿, 加载在样式中,通过样式运用于单元格
     * 背景色
     * @throws IOException
     */
    public void testExcelStyle() throws IOException {
        // 1. 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 1.1 创建单元格对象 合并第三行第三列到5列
        // 构造参数 起始行号 结束行号 起始列号 结束列号
        CellRangeAddress cellRangeAddress = new CellRangeAddress(
                2, 2, 2, 4 );
        // 1.2 创建单元格样式
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setAlignment(HSSFCellStyle.VERTICAL_CENTER);

        // 1.3 创建字体
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints((short)16);
        // 将字体加载到样式中
        style.setFont(font);

        // 1.4 设置背景色为黄色
        // 1.4.1 设置填充模式
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(HSSFColor.YELLOW.index);
        style.setFillForegroundColor(HSSFColor.GREEN.index);

        // 2. 创建工作类
        HSSFSheet sheet = workbook.createSheet("hello world");
        // 2.1 加入合并单元格对象
        sheet.addMergedRegion(cellRangeAddress);

        // 3. 创建行 , 第三行 注意:从0开始
        HSSFRow row = sheet.createRow(2);

        // 4. 创建单元格, 第三行第三列 注意:从0开始
        HSSFCell cell = row.createCell(2);
        cell.setCellValue("Hello World");
        // 4.1 单元格添加样式
        cell.setCellStyle(style);

        String fileName = "D:\\huangMP\\Desktop\\HelloExcelStyle.xls";
        FileOutputStream fileOutputSteam = new FileOutputStream(fileName);

        workbook.write(fileOutputSteam);
        workbook.close();

        fileOutputSteam.close();

    }
}
