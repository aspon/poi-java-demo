package cn.huangmp.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by huangMP on 2017/8/20.
 * decription :
 */
public class OutputExcelDemo {

    /**
     *  从 工作簿中写入 数据 OutputExcelDemo
     * 07 版本及之前的版本写法
     * @throws IOException
     */
    public void outputExcel() throws IOException {
        // 1. 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 2. 创建工作类
        HSSFSheet sheet = workbook.createSheet("hello world");

        // 3. 创建行 , 第三行 注意:从0开始
        HSSFRow row = sheet.createRow(2);

        // 4. 创建单元格, 第三行第三列 注意:从0开始
        HSSFCell cell = row.createCell(2);
        cell.setCellValue("Hello World");

        String fileName = "D:\\huangMP\\Desktop\\OutputExcelDemo.xls";
        FileOutputStream fileOutputSteam = new FileOutputStream(fileName);

        workbook.write(fileOutputSteam);
        workbook.close();

        fileOutputSteam.close();
    }
}
