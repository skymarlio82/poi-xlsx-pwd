package com.mylin.controller;

import com.mylin.model.ExcelData;
import com.mylin.util.ExcelUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

@RestController
public class ExcelController {

    @RequestMapping(value="/excel", method=RequestMethod.GET)
    public void excel(HttpServletResponse response) throws Exception {
        ExcelData data = new ExcelData();
        data.setName("hello");
        List<String> titles = new ArrayList();
        titles.add("a1");
        titles.add("a2");
        titles.add("a3");
        data.setTitles(titles);
        List<List<Object>> rows = new ArrayList();
        List<Object> row = new ArrayList();
        row.add("11111111111");
        row.add("22222222222");
        row.add("33333333333");
        rows.add(row);
        data.setRows(rows);
        ExcelUtil.exportExcel(data);
        ExcelUtil.encryptExcel("password");
        ExcelUtil.exportExcel(response, "hello.xlsx");
    }
}