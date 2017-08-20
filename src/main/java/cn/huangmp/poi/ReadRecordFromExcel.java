package cn.huangmp.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by huangMP on 2017/8/20.
 * decription :
 */
public class ReadRecordFromExcel {


    /**
     * 从 工作簿中读取 数据 OutputExcelDemo
     * @throws IOException
     */
    public void readExel() throws IOException, InvalidFormatException {

        String fileName = "D:\\huangMP\\Desktop\\OutputRecordFromExcel.xls";
        FileInputStream fileInputStream = new FileInputStream(fileName);

        // 1. 读取工作簿
        Workbook workbook = WorkbookFactory.create(fileInputStream);

        // 2. 读取工作表
        Sheet sheet = workbook.getSheetAt(0);

        List<UserInfo> users = null;
        // 3. 读取行
        if(sheet.getPhysicalNumberOfRows()>2){
            users = new ArrayList<UserInfo>();
            // 4. 读取单元格
            for(int i = 2 ; i < sheet.getPhysicalNumberOfRows() ; i++ ){
                Row row = sheet.getRow(i);
                UserInfo userInfo = new UserInfo();
                userInfo.setName(row.getCell(0).getStringCellValue());
                userInfo.setUsername(row.getCell(1).getStringCellValue());
                userInfo.setDepartment(row.getCell(2).getStringCellValue());
                userInfo.setSex((int)row.getCell(3).getNumericCellValue());
                userInfo.setEmail(row.getCell(4).getStringCellValue());
                users.add(userInfo);
            }
        }

        System.out.println("读取到的人数为 : " + users.size());

        workbook.close();
        fileInputStream.close();
    }
}
