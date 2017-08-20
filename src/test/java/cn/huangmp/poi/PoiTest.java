package cn.huangmp.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by huangMP on 2017/8/20.
 * decription :
 */
public class PoiTest {
    @Test
    public void outputExcel() throws Exception {
        OutputExcelDemo outputExcelDemo = new OutputExcelDemo();
        outputExcelDemo.outputExcel();
    }

    @Test
    public void readExcel() throws Exception {
        ReadExcelDemo readExel = new ReadExcelDemo();
        readExel.readExel();
    }


    @Test
    public void testOutputUserInfoToExcel() throws IOException {
        List<UserInfo> users = new ArrayList<UserInfo>();
        // 创建测试数据
        for(int i = 0 ; i < 50 ; i++ ){
            UserInfo u = new UserInfo();
            u.setName(String.valueOf(i));
            u.setUsername(String.valueOf(i));
            u.setDepartment(String.valueOf(i));
            u.setSex(i);
            u.setEmail(String.valueOf(i));
            users.add(u);
        }

        OutputUserInfoRecord outputUserInfoRecord = new OutputUserInfoRecord();
        outputUserInfoRecord.outputRecordFromExcel(users);
    }

    @Test
    public void testReadUserInfoRecordFromExcel() throws IOException, InvalidFormatException {
        ReadRecordFromExcel readRecordFromExcel = new ReadRecordFromExcel();
        readRecordFromExcel.readExel();
    }
}