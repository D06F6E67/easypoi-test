package cn.afterturn.easypoi.test.excel.template;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.test.entity.CourseEntity;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;
import org.junit.Test;

import java.io.FileOutputStream;
import java.util.*;

import static cn.afterturn.easypoi.excel.ExcelExportUtil.SHEET_NAME;

/**
 * @author by jueyue on 19-6-11.
 */
public class ExcelExportTemplateCloneTest {

    List<CourseEntity> list = new ArrayList<>();

    @Test
    public void test1() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "test/demos.xlsx", true);
        params.setHeadingRows(4);


        List<Map<String, Object>> numOneList = new ArrayList<>();

        for (int k = 0; k < 5; k++) {
            Map<String, Object> map = new HashMap<>();
            List<Map<String, String>> listMap = new ArrayList<>();
            for (int i = 1; i <= 10; i++) {
                Map<String, String> lm = new HashMap<>();
                for (int j = 1; j <= 48; j++) {
                    boolean line = ((int) (Math.random() * 100)) % 2 == 0;
                    lm.put(String.valueOf(j), line ? "ONLINE" : "OFFLINE");
                }
                listMap.add(lm);
            }
            map.put("list", listMap);

            map.put(SHEET_NAME, k);

            numOneList.add(map);
        }

        Map<Integer, List<Map<String, Object>>> realMap = new HashMap<>();
        realMap.put(0, numOneList);
//        realMap.put(1, numOneList);

        XSSFWorkbook book = (XSSFWorkbook) ExcelExportUtil.exportExcelClone(realMap, params);
        book.setForceFormulaRecalculation(true);
        FileOutputStream fos = new FileOutputStream("/Users/lee/Downloads/test.xlsx");
        book.write(fos);
        fos.close();
    }

    @Test
    public void test2() throws Exception {
        Map<String, Object> value = new HashMap<>();

        List<Map<String, Object>> colList = new ArrayList<>();
        // 先处理表头
        Map<String, Object> map = new HashMap<>();
        map.put("name", "2024-01-01");
        map.put("q", "Quantity");
        map.put("a", "Amount");
        map.put("qv", "t.qv_01");
        map.put("av", "t.av_01");
        colList.add(map);

        map = new HashMap<>();
        map.put("name", "2024-01-02");
        map.put("q", "Quantity");
        map.put("a", "Amount");
        map.put("qv", "t.qv_02");
        map.put("av", "t.av_02");
        colList.add(map);

        map = new HashMap<>();
        map.put("name", "2024-01-03");
        map.put("q", "Quantity");
        map.put("a", "Amount");
        map.put("qv", "t.qv_03");
        map.put("av", "t.av_03");
        colList.add(map);
        value.put("colList", colList);


        List<Map<String, Object>> valList = new ArrayList<>();
        map = new HashMap<>();
        map.put("one", "菜单1");
        map.put("two", "菜品1");
        map.put("qv_01", 1);
        map.put("av_01", 2);
        map.put("qv_02", 3);
        map.put("av_02", 4);
        valList.add(map);
        map = new HashMap<>();
        map.put("one", "菜单1");
        map.put("two", "菜品2");
        map.put("qv_01", 1);
        map.put("av_01", 2);
        map.put("qv_02", 3);
        map.put("av_02", 4);
        valList.add(map);
        map = new HashMap<>();
        map.put("one", "菜单2");
        map.put("two", "菜品3");
        map.put("qv_01", 1);
        map.put("av_01", 2);
        map.put("qv_02", 3);
        map.put("av_02", 4);
        valList.add(map);

        map = new HashMap<>();
        map.put("one", "Total");
        map.put("qv_01", 5);
        map.put("av_01", 6);
        map.put("qv_02", 7);
        map.put("av_02", 8);
        valList.add(map);
        value.put("valList", valList);


        TemplateExportParams params = new TemplateExportParams(
                "test/demo2.xlsx");
        params.setColForEach(true);
        Workbook book = ExcelExportUtil.exportExcel(params, value);
        FileOutputStream fos = new FileOutputStream("/Users/lee/Downloads/test.xlsx");
        book.write(fos);
        fos.close();
    }

    @Test
    public void cloneTest() throws Exception {
        TemplateExportParams params = new TemplateExportParams(
                "doc/exportTemp.xls", true);
        params.setHeadingRows(2);
        params.setHeadingStartRow(2);
        // params.setStyle(ExcelStyleType.BORDER.getClazz());
        List<Map<String, Object>> numOneList = new ArrayList<>();
        for (int i = 0; i < 3; i++) {
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
        for (int i = 0; i < 3; i++) {
            Map<String, Object> oneMap = new HashMap<>();
            oneMap.put("list", list);
            oneMap.put(SHEET_NAME, "第二个测试SHeet" + i);
            numTowList.add(oneMap);
        }


        Map<Integer, List<Map<String, Object>>> realMap = new HashMap<>();
        // realMap.put(0, numOneList);
        realMap.put(1, numOneList);

        Workbook book = ExcelExportUtil.exportExcelClone(realMap, params);

        FileOutputStream fos = new FileOutputStream("/Users/lee/Downloads/exportCloneTemp.xls");
        book.write(fos);
        fos.close();

    }
}