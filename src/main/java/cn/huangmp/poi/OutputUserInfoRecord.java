package cn.huangmp.poi;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * Created by huangMP on 2017/8/20.
 * decription :
 */
public class OutputUserInfoRecord {


    /**
     * 导出用户列表
     */
    public void outputRecordFromExcel(List<UserInfo> users) throws IOException {

        // 1. 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 1.1 创建合并单元格对象
        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 4);
        // 1.2 创建头标题行样式并创建字体
        HSSFCellStyle style1 = createCellStyle(workbook, (short)16);

        // 1.3 创建标题样式
        HSSFCellStyle style2 = createCellStyle(workbook, (short)13);

        // 2. 创建工作表
        HSSFSheet sheet = workbook.createSheet("用户列表");

        // 2.1 加载合并单元格对象
        sheet.addMergedRegion(cellRangeAddress);
        // 2.2 设置默认列宽
        sheet.setDefaultColumnWidth(20);

        // 3. 创建行
        // 3.1 创建头标题行并写入头标题
        HSSFRow row1 = sheet.createRow(0);
        HSSFCell cell1 = row1.createCell(0);
        cell1.setCellStyle(style1);
        cell1.setCellValue("用户列表");
        // 3.2 创建列标题并写入列标题
        HSSFRow row2 = sheet.createRow(1);
        String[] titles = {"用户名称", "账号", "所属部门", "性别", "邮箱"};
        for (int i = 0 ; i < titles.length ; i++ ) {
            HSSFCell cell2 = row2.createCell(i);
            cell2.setCellStyle(style2);
            cell2.setCellValue(titles[i]);
        }

        // 4. 创建单元格,写入用户数据到excel
        if(users != null && users.size() > 0){
            for(int j = 0 ; j < users.size(); j++ ){
                HSSFRow row = sheet.createRow(j+2);
                row.createCell(0).setCellValue(users.get(j).getName());
                row.createCell(1).setCellValue(users.get(j).getUsername());
                row.createCell(2).setCellValue(users.get(j).getDepartment());
                row.createCell(3).setCellValue(users.get(j).getSex());
                row.createCell(4).setCellValue(users.get(j).getEmail());
            }
        }

        // 5. 输出
        String fileName = "D:\\huangMP\\Desktop\\OutputRecordFromExcel.xls";
        FileOutputStream fileOutputSteam = new FileOutputStream(fileName);

        workbook.write(fileOutputSteam);
        workbook.close();

        fileOutputSteam.close();
    }

    private HSSFCellStyle createCellStyle(HSSFWorkbook workbook, short fontSize){
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        // 创建字体
        HSSFFont hssfFont = workbook.createFont();
        hssfFont.setFontHeightInPoints((short)fontSize);
        // 在样式中加载字体
        style.setFont(hssfFont);
        return style;
    }
}
