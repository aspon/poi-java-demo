package cn.huangmp.poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by huangMP on 2017/8/20.
 * decription :
 */
public class ReadExcelDemo {
    /**
     * 从 工作簿中读取 数据 ReadExcelDemo
     * @throws IOException
     */
    public void readExel() throws IOException {

        String fileName = "D:\\huangMP\\Desktop\\OutputExcelDemo.xls";
        FileInputStream fileInputStream = new FileInputStream(fileName);

        // 1. 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);

        // 2. 创建工作类
        HSSFSheet sheet = workbook.getSheetAt(0);

        // 3. 创建行 , 第三行 注意:从0开始
        HSSFRow row = sheet.getRow(2);

        // 4. 创建单元格, 第三行第三列 注意:从0开始
        HSSFCell cell = row.getCell(2);
        String cellString = cell.getStringCellValue();

        System.out.println("第三行第三列的值为 : " + cellString );

        workbook.close();
        fileInputStream.close();
    }
}
