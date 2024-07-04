package cn.afterturn.easypoi.test.excel.template;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.test.entity.CourseEntity;
import cn.afterturn.easypoi.test.entity.StudentEntity;
import cn.afterturn.easypoi.test.entity.TeacherEntity;
import cn.afterturn.easypoi.util.PoiMergeCellUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Before;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.*;

import static cn.afterturn.easypoi.excel.ExcelExportUtil.SHEET_NAME;

/**
 * @author by jueyue on 19-6-11.
 */
public class ExcelExportTemplateClone {

    List<CourseEntity> list = new ArrayList<>();
    CourseEntity courseEntity;

    @Test
    public void test1() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "test/adv.xlsx", true);
        params.setHeadingRows(12);


        List<Map<String, Object>> numOneList = new ArrayList<>();

        for (int k = 0; k < 5; k++) {
            Map<String, Object> map = new HashMap<>();
            List<Map<String, String>> listMap = new ArrayList<>();
            for (int i = 1; i <= 10; i++) {
                int sum = 0;
                Map<String, String> lm = new HashMap<>();
                lm.put("date", String.format("5-%s", i));
                lm.put("day", String.format("day %s", i));
                lm.put("a", "200");
                for (int j = 0; j < 24; j++) {
                    int num = (int) (Math.random() * 100);
                    sum += num;
                    lm.put(String.valueOf(j), String.valueOf(num));
                }
                lm.put("b", String.valueOf(sum));
                listMap.add(lm);
            }
            map.put("maplist", listMap);

            if (k == 0) {
                map.put(SHEET_NAME, "SS");
            } else {

                map.put(SHEET_NAME, "S" + k);
            }
            numOneList.add(map);
        }

        List<Map<String, Object>> numTwoList = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put(SHEET_NAME, "MAIN");
        numTwoList.add(map);

        Map<Integer, List<Map<String, Object>>> realMap = new HashMap<>();
        realMap.put(0, numTwoList);
        realMap.put(1, numOneList);

        XSSFWorkbook book = (XSSFWorkbook) ExcelExportUtil.exportExcelClone(realMap, params);

        XSSFHyperlink hyperlink = book.getSheetAt(0).getRow(4).getCell(1).getHyperlink();
        hyperlink.setLocation("'SS'!A1");
        book.getSheetAt(0).getRow(4).getCell(1).setHyperlink(hyperlink);

        XSSFDrawing d0 = book.getSheet("SS").getDrawingPatriarch();

        XSSFPicture xssfPicture = (XSSFPicture) (d0.getShapes().get(1));

        XSSFShape sh0 = d0.getShapes().get(0);
        XSSFClientAnchor anchor = (XSSFClientAnchor) sh0.getAnchor();
        book.setForceFormulaRecalculation(true);
        // Workbook book = ExcelExportUtil.exportExcel(params, map);
        FileOutputStream fos = new FileOutputStream("/Users/lee/Downloads/test.xlsx");
        book.write(fos);
        fos.close();
    }

    @Test
    public void xmeta() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "test/xmeta.xlsx", true);

        // sheet1数据
        List<Map<String, Object>> numList1 = new ArrayList<>();
        Map<String, Object> map1 = new HashMap<>();
        List<Map<String, String>> listMap1 = new ArrayList<>();
        for (int i = 1; i < 4; i++) {
            Map<String, String> lm = new HashMap<>();
            lm.put("date", "2024.04.0" + i);
            lm.put("quantity", "100" + i);
            lm.put("amount", "20000" + i);
            lm.put("payAmount", "180000" + i);
            lm.put("xmetaAmount", "170000" + i);

            lm.put("ipayQuantity", "20" + i);
            lm.put("ipayAmount", "50000" + i);
            lm.put("ipayPayAmount", "45000" + i);
            lm.put("ipayXmetaAmount", "40500" + i);

            lm.put("gcashQuantity", "100" + i);
            lm.put("gcashAmount", "55500" + i);
            lm.put("gcashPayAmount", "1000" + i);
            lm.put("gcashXmetaAmount", "10000" + i);

            lm.put("paymayaQuantity", "100" + i);
            lm.put("paymayaAmount", "100" + i);
            lm.put("paymayaPayAmount", "100" + i);
            lm.put("paymayaXmetaAmount", "100" + i);

            lm.put("onlineBankQuantity", "100" + i);
            lm.put("onlineBankAmount", "100" + i);
            lm.put("onlineBankPayAmount", "100" + i);
            lm.put("onlineBankXmetaAmount", "100" + i);

            lm.put("bankQuantity", "100" + i);
            lm.put("bankAmount", "100" + i);
            lm.put("bankPayAmount", "100" + i);
            lm.put("bankXmetaAmount", "100" + i);

            lm.put("cashQuantity", "100" + i);
            lm.put("cashAmount", "100" + i);
            lm.put("cashPayAmount", "100" + i);
            lm.put("cashXmetaAmount", "100" + i);

            lm.put("tripartiteQuantity", "100" + i);
            lm.put("tripartiteAmount", "100" + i);
            lm.put("tripartitePayAmount", "100" + i);
            lm.put("tripartiteXmetaAmount", "100" + i);

            lm.put("grabfoodQuantity", "100" + i);
            lm.put("grabfoodAmount", "100" + i);
            lm.put("grabfoodPayAmount", "100" + i);
            lm.put("grabfoodXmetaAmount", "100" + i);

            lm.put("foodPandaQuantity", "100" + i);
            lm.put("foodPandaAmount", "100" + i);
            lm.put("foodPandaPayAmount", "100" + i);
            lm.put("foodPandaXmetaAmount", "100" + i);

            lm.put("otherQuantity", "100" + i);
            lm.put("otherAmount", "100" + i);
            lm.put("otherPayAmount", "100" + i);
            lm.put("otherXmetaAmount", "100" + i);

            lm.put("memberConsumptionQuantity", "100" + i);
            lm.put("memberConsumptionAmount", "100" + i);
            lm.put("memberConsumptionPayAmount", "100" + i);
            lm.put("memberConsumptionXmetaAmount", "100" + i);

            lm.put("memberRechargeQuantity", "100" + i);
            lm.put("memberRechargeAmount", "100" + i);
            lm.put("memberRechargePayAmount", "100" + i);
            lm.put("memberRechargeXmetaAmount", "100" + i);

            listMap1.add(lm);
        }
        map1.put("maplist", listMap1);
        map1.put(SHEET_NAME, "销售汇总");
        numList1.add(map1);

        // sheet2数据
        List<Map<String, Object>> numList2 = new ArrayList<>();
        Map<String, Object> map2 = new HashMap<>();
        List<Map<String, String>> listMap2 = new ArrayList<>();
        for (int i = 1; i < 4; i++) {
            Map<String, String> lm = new HashMap<>();
            lm.put("date", "2024.04.0" + i);
            lm.put("orderNumber", "1774737110628610048" + i);
            lm.put("payType", "Direct Merchant Recharge" + i);
            lm.put("orderType", "会员卡充值订单" + i);

            lm.put("orderAmount", "99" + i);
            lm.put("payAmount", "99" + i);
            lm.put("actualAmount", "99" + i);
            lm.put("tip", "9" + i);
            lm.put("tax", "9" + i);
            lm.put("deliveryFee", "99." + i);
            lm.put("discountedPrice", "9" + i);
            lm.put("offerType", "促销" + i);
            lm.put("refundAmount", "9" + i);

            listMap2.add(lm);
        }
        map2.put("maplist", listMap2);
        map2.put(SHEET_NAME, "销售订单详情");
        numList2.add(map2);

        // sheet3数据
        List<Map<String, Object>> numList3 = new ArrayList<>();
        Map<String, Object> map3 = new HashMap<>();
        List<Map<String, String>> listMap3 = new ArrayList<>();
        for (int i = 1; i < 6; i++) {
            Map<String, String> lm = new HashMap<>();
            lm.put("date", "2024.04.0" + i/4);
            lm.put("name", "酱爆茄子" + i/3);
            lm.put("type", "小炒时蔬" + i/2);
            lm.put("quantity", "1" + i);
            lm.put("amount", "10" + i);

            listMap3.add(lm);
        }
        map3.put("maplist", listMap3);
        map3.put(SHEET_NAME, "商品销售详情");
        numList3.add(map3);

        // sheet4数据
        List<Map<String, Object>> numList4 = new ArrayList<>();
        Map<String, Object> map4 = new HashMap<>();
        List<Map<String, String>> listMap4 = new ArrayList<>();
        for (int i = 1; i < 4; i++) {
            Map<String, String> lm = new HashMap<>();
            lm.put("code", "DB00001" + i);
            lm.put("name", "ZHOU" + i);
            lm.put("tel", "99999" + i);
            lm.put("birthday", "1999/9/" + i);
            lm.put("grade", "" + i);
            lm.put("balance", "200" + i);
            lm.put("rechargeAmount", "2000" + i);
            lm.put("log", "2024/1/1 +10000\n" +
                    "2024/1/3 +8000\n" +
                    "2024/1/3 +8000");

            listMap4.add(lm);
        }
        map4.put("maplist", listMap4);
        map4.put(SHEET_NAME, "会员卡明细");
        numList4.add(map4);

        Map<Integer, List<Map<String, Object>>> realMap = new HashMap<>();
        realMap.put(0, numList1);
        realMap.put(1, numList2);
        realMap.put(2, numList3);
        realMap.put(3, numList4);
        XSSFWorkbook book = (XSSFWorkbook) ExcelExportUtil.exportExcelClone(realMap, params);

        FileOutputStream fos = new FileOutputStream("/Users/lee/Downloads/test.xlsx");
        book.write(fos);
    }

    @Test
    public void mergeCol() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "test/mergeCol.xlsx", true);

        List<Map<String, Object>> numList3 = new ArrayList<>();
        Map<String, Object> map3 = new HashMap<>();
        List<Map<String, String>> listMap3 = new ArrayList<>();
        for (int i = 1; i < 8; i++) {
            Map<String, String> lm = new HashMap<>();
            lm.put("date", "2024.04.0" + i/4);
            lm.put("name", "酱爆茄子" + i/3);
            lm.put("type", "小炒时蔬" + i/2);
            lm.put("quantity", "1" + i);
            lm.put("amount", "10" + i);

            listMap3.add(lm);
        }
        map3.put("maplist", listMap3);
        map3.put(SHEET_NAME, "商品销售详情");
        numList3.add(map3);

        Map<Integer, List<Map<String, Object>>> realMap = new HashMap<>();
        realMap.put(0, numList3);
        XSSFWorkbook book = (XSSFWorkbook) ExcelExportUtil.exportExcelClone(realMap, params);

        // Sheet sheet = book.getSheet("商品销售详情");
        // Map<Integer, int[]> mergeMap = new HashMap<>();
        // mergeMap.put(0, null);
        // mergeMap.put(1, new int[]{0});
        // mergeMap.put(2, new int[]{1});
        // PoiMergeCellUtil.mergeCells(sheet, mergeMap,1);

        FileOutputStream fos = new FileOutputStream("/Users/lee/Downloads/test.xlsx");
        book.write(fos);
    }

    @Test
    public void cloneTest() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "doc/exportTemp.xls", true);
        params.setHeadingRows(2);
        params.setHeadingStartRow(2);
        List<Map<String, Object>> numOneList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> map = new HashMap<String, Object>();
            // sheet 1
            map.put("year", "2013" + i);
            map.put("sunCourses", list.size());
            Map<String, Object> obj = new HashMap<String, Object>();
            map.put("obj", obj);
            obj.put("name", list.size());
            // sheet 2
            map.put("month", 10);
            Map<String, Object> temp;
            for (int j = 1; j < 8; j++) {
                temp = new HashMap<String, Object>();
                temp.put("per", j * 10 + "---" + i);
                temp.put("mon", j * 1000);
                temp.put("summon", j * 10000);
                map.put("i" + j, temp);
            }
            map.put(SHEET_NAME, "啊啊测试SHeet" + i);
            numOneList.add(map);
        }


        List<Map<String, Object>> numTowList = new ArrayList<>();
        for (int i = 0; i < 15; i++) {
            Map<String, Object> oneMap = new HashMap<String, Object>();
            oneMap.put("list", list);
            oneMap.put(SHEET_NAME, "第二个测试SHeet" + i);
            numTowList.add(oneMap);
        }


        Map<Integer, List<Map<String, Object>>> realMap = new HashMap<>();
        realMap.put(0, numOneList);
        realMap.put(1, numTowList);

        Workbook book = ExcelExportUtil.exportExcelClone(realMap, params);
        // File     savefile = new File("D:/home/excel/");
        // if (!savefile.exists()) {
        //     savefile.mkdirs();
        // }
        FileOutputStream fos = new FileOutputStream("/Users/lee/Downloads/exportCloneTemp.xls");
        book.write(fos);
        fos.close();

    }

    @Before
    public void testBefore() {
        courseEntity = new CourseEntity();
        courseEntity.setId("1131");
        courseEntity.setName("小白");

        TeacherEntity teacherEntity = new TeacherEntity();
        teacherEntity.setId("12131231");
        teacherEntity.setName("你们");
        courseEntity.setChineseTeacher(teacherEntity);

        teacherEntity = new TeacherEntity();
        teacherEntity.setId("121312314312421131");
        teacherEntity.setName("老王");
        courseEntity.setMathTeacher(teacherEntity);

        StudentEntity studentEntity = new StudentEntity();
        studentEntity.setId("11231");
        studentEntity.setName("撒旦法司法局");
        studentEntity.setBirthday(new Date());
        studentEntity.setSex(1);
        List<StudentEntity> studentList = new ArrayList<StudentEntity>();
        studentList.add(studentEntity);
        studentList.add(studentEntity);
        courseEntity.setStudents(studentList);

        for (int i = 0; i < 3; i++) {
            list.add(courseEntity);
        }
    }
}