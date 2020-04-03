package com.kovizone.poi.ooxml.plus.test;

import com.kovizone.poi.ooxml.plus.WorkbookCreator;
import com.kovizone.poi.ooxml.plus.anno.*;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

public class Test {


    public static void main(String[] args) throws IOException, PoiOoxmlPlusException {
        List<TestEntity> testEntities = new ArrayList<>();
        testEntities.add(new TestEntity("1", "2", "00", "4", 1, new Date(), null));
        testEntities.add(new TestEntity("5", "6", "01", "8", 2, new Date(), "test"));
        testEntities.add(new TestEntity("9", "10", "02", null, 3, new Date(), null));

        Workbook workbook = new HSSFWorkbook();
        new WorkbookCreator().constructor(workbook, testEntities, new HashMap<String, String>() {{
            put("#{datetime}", new SimpleDateFormat("yyyy年MM月dd日 HH:mm:ss").format(new Date()));
            put("#{author}", "KoviChen");
        }});
        workbook.write(new FileOutputStream(new File("D:/test/" + new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date()) + ".xls")));
    }

    @SheetConfig(sheetName = "测试")
    @HeaderRender(headerTitle = "这是标题", headerSubtitle = "日期：#{datetime} 作者：#{author}")
    static class TestEntity {
        @Prefix("#[test2]")
        @Suffix("后缀")
        @ColumnConfig(sort = 10, title = "测试1", width = 5000)
        String test;

        @ColumnConfig(sort = 20, title = "测试2")
        String test2;

        @StringReplace(target = {"00", "01", "02"}, replacement = {"交易完成", "交易中", "交易失败"})
        @ColumnConfig(sort = 30, title = "测试3")
        String test3;

        @StringReplace(target = {"00", "01", "02"}, replacement = {"交易完成", "交易中", "交易失败"})
        @Substitute(condition = "is empty", value = "#[test3]")
        @ColumnConfig(sort = 40, title = "测试4")
        String test4;

        @NumberMapper(target = {1, 2, 3}, replacement = {"交易完成", "交易中", "交易失败"})
        @ColumnConfig(sort = 50, title = "测试5")
        Integer test5;

        @Criteria("#[test4] is null")
        @DateFormat("yyyy-MM-dd HH:mm:ss.SSS")
        @ColumnConfig(sort = 60, title = "测试6", width = 10000)
        Date test6;

        /**
         * PoiDefaultValue可以在该属性为空时指定一个默认值
         */
        @DefaultValue("默认值")
        @ColumnConfig(sort = 70, title = "测试7", width = 10000)
        String test7;

        public TestEntity(String test, String test2, String test3, String test4, Integer test5, Date test6, String test7) {
            this.test = test;
            this.test2 = test2;
            this.test3 = test3;
            this.test4 = test4;
            this.test5 = test5;
            this.test6 = test6;
            this.test7 = test7;
        }

        public String getTest7() {
            return test7;
        }

        public void setTest7(String test7) {
            this.test7 = test7;
        }

        public TestEntity() {

        }

        public Date getTest6() {
            return test6;
        }

        public void setTest6(Date test6) {
            this.test6 = test6;
        }

        public Integer getTest5() {
            return test5;
        }

        public void setTest5(Integer test5) {
            this.test5 = test5;
        }

        public String getTest() {
            return test;
        }

        public void setTest(String test) {
            this.test = test;
        }

        public String getTest2() {
            return test2;
        }

        public void setTest2(String test2) {
            this.test2 = test2;
        }

        public String getTest3() {
            return test3;
        }

        public void setTest3(String test3) {
            this.test3 = test3;
        }

        public String getTest4() {
            return test4;
        }

        public void setTest4(String test4) {
            this.test4 = test4;
        }
    }
}
