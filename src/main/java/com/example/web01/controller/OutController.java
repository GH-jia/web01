package com.example.web01.controller;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.example.web01.utils.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.util.IOUtils;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.client.RestTemplate;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.ByteBuffer;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;


@RestController
@RequestMapping("/out")
public class OutController {

    /**
     * 获取本地资源输出到前端 excel
     */
    @GetMapping("getXls")
    public String getXls(HttpServletResponse response){
        //第一行
        String[] title = {"部门","工号","姓名","性别"};
        //生成表格对象
        HSSFWorkbook hssfWorkbook = ExcelUtil.getHSSFWorkbook("东大", title, new HSSFWorkbook());
        //获取sheet
        HSSFSheet sheet = hssfWorkbook.getSheet("东大");

        try{
            //获取本地文件
            File f = new ClassPathResource("files/aa.json").getFile();
            if (!f.exists()){
                return null;
            }
            String s = null;
            //从输入流获取字节数组并转为字符串
            FileInputStream fileInputStream = new FileInputStream(f);
            BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
            //available 方法似乎不靠谱，只能获取一个估计值
            int len = bufferedInputStream.available();
            byte[] bs = new byte[len];
            bufferedInputStream.read(bs);
            bufferedInputStream.close();
            s = new String(bs, StandardCharsets.UTF_8);
            System.out.println(s);
            //字符串转为json对象,获取想要的数据
            JSONObject jsonObject = JSON.parseObject(s);
            JSONArray rows = jsonObject.getJSONArray("rows");
            int n = 0;
            for (int i=0;i< rows.size();i++){
                JSONObject j1 = rows.getJSONObject(i);
                HSSFRow row = sheet.createRow(++n);
                for (int l=0;l< title.length;++l){
                    if (0 == l){
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue(j1.getJSONObject("dept").getString("deptName"));
                        System.out.println(j1.getJSONObject("dept").getString("deptName"));
                    }else if (1 == l){
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue(j1.getString("gh"));
                        System.out.println(j1.getString("gh"));
                    }else if (2 == l){
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue(j1.getString("xm"));
                    }else if(3 == l) {
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue((j1.getIntValue("sex") == 0 ? "男":"女"));
                    }
                }
            }
        }catch (Exception e){
            e.printStackTrace();
        }

        //输出excle文件
        try{
            // 设置下载时客户端Excel的名称
            SimpleDateFormat simpleDateFormat1 = new SimpleDateFormat("yyyy-MM-dd");
            String format = simpleDateFormat1.format(new Date());
            String fileName = "dd"+format+".xls";
            //文件编码
            String encodedFileName = URLEncoder.encode(fileName, "UTF-8");

            response.setContentType("application/vnd.ms-excel");
            // 解决中文乱码
            response.setHeader("Content-Disposition", "attachment;");
            response.setHeader("Content-Disposition", "attachment;filename=\"" +encodedFileName+"\"");
            ServletOutputStream out = response.getOutputStream();
            hssfWorkbook.write(out);
            out.flush();
            out.close();
        }catch (Exception e){

        }
        return "success";
    }

    //http获取数据，输出文件到本地 excel
    @GetMapping("/getXlsByNet")
    public String getXlsByNet(){
        //第一行
        String[] title = {"部门","工号","姓名","性别"};
        //生成表格对象
        HSSFWorkbook hssfWorkbook = ExcelUtil.getHSSFWorkbook("东大", title, new HSSFWorkbook());
        //获取sheet
        HSSFSheet sheet = hssfWorkbook.getSheet("东大");
        int n = 0;
        for (int j=0;j<=(950/20);j++){
            //通过链接拿到字符串数据
            int start = 20;
            start = start*j;
            String s = getData(String.valueOf(start));
            JSONObject jsonObject = JSON.parseObject(s);
            JSONArray rows = jsonObject.getJSONArray("rows");
            //解析json数据
            for (int i=0;i< rows.size();i++){
                JSONObject j1 = rows.getJSONObject(i);
                HSSFRow row = sheet.createRow(++n);
                for (int l=0;l< title.length;++l){
                    if (0 == l){
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue(j1.getJSONObject("dept").getString("deptName"));
                    }else if (1 == l){
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue(j1.getString("gh"));
                    }else if (2 == l){
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue(j1.getString("xm"));
                    }else if (3 == l){
                        HSSFCell cell = row.createCell(l);
                        cell.setCellValue((j1.getIntValue("sex") == 0 ? "男":"女"));
                    }
                }
            }
        }

        //输出excle文件
        try{
            BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream("D:\\tmp\\dd.xls"));
            hssfWorkbook.write(out);
            out.flush();
            out.close();
        }catch (Exception e){

        }

        return "";
    }
    private String getData(String start){
        //通过链接拿到字符串数据
        RestTemplate restTemplate = new RestTemplate();
        String url  = "http://oa.neuq.edu.cn/hr/teacherInfo/queryListForPage";
        HttpHeaders headers = new HttpHeaders();
        MediaType type = MediaType.parseMediaType("application/x-www-form-urlencoded; charset=UTF-8");
        headers.setContentType(type);
        MultiValueMap<String, String> map = new LinkedMultiValueMap<>();
        map.add("start", start);
        map.add("limit", "20");
        map.add("field", "gh");
        HttpEntity<MultiValueMap<String, String>> httpEntity = new HttpEntity<>(map, headers);
        String result = restTemplate.postForObject(url, httpEntity, String.class);
        System.out.println(result);
        return result;
    }

    /**
     * 获取本地资源输出到前端 wrod
     */
    @GetMapping("/getDocx")
    public void getDocx(HttpServletResponse response){

        try{
            //中文文件名有可能会导致jar中读取不到文件
            InputStream inputStream = new ClassPathResource("files/wordm.docx").getInputStream();
            response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.setCharacterEncoding("UTF-8");
            ServletOutputStream outputStream = response.getOutputStream();
            IOUtils.copy(inputStream,outputStream);
            inputStream.close();
            outputStream.flush();
            outputStream.close();
        }catch (Exception e){
            StringWriter stringWriter = new StringWriter();
            PrintWriter printWriter = new PrintWriter(stringWriter);
            e.printStackTrace(printWriter);
            System.out.println(stringWriter.toString());
        }
    }

}
