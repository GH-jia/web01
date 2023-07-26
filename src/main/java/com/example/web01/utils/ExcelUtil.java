package com.example.web01.utils;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.usermodel.CellType.*;

/**
 * Excel工具类
 *
 * @author Administrator
 */
public class ExcelUtil {
    private  static POIFSFileSystem fs;//poi文件流
    private  static Workbook wb;//获得execl
    private  static Row row;//获得行
    private  static Sheet sheet;//获得工作簿

    /**
     * 导出Excel
     *
     * @param sheetName sheet名称
     * @param title     标题
     * @param values    内容
     * @param wb        HSSFWorkbook对象
     * @return HSSFWorkbook
     */
    public static HSSFWorkbook getHSSFWorkbook(String sheetName, String[] title, String[][] values, HSSFWorkbook wb) {

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        if (wb == null) {
            wb = new HSSFWorkbook();
        }

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        // 创建一个居中格式
        style.setAlignment(HorizontalAlignment.CENTER);

        //声明列对象
        HSSFCell cell = null;

        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        //创建内容
        for (int i = 0; i < values.length; i++) {
            row = sheet.createRow(i + 1);
            for (int j = 0; j < values[i].length; j++) {
                //将内容按顺序赋给对应的列对象
                row.createCell(j).setCellValue(values[i][j]);
            }
        }
        return wb;
    }
    /**
     * 导出Excel
     *
     * @param sheetName sheet名称
     * @param title     标题
     * @param wb        HSSFWorkbook对象
     * @return HSSFWorkbook
     */
    public static HSSFWorkbook getHSSFWorkbook(String sheetName, String[] title, HSSFWorkbook wb) {

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        if (wb == null) {
            wb = new HSSFWorkbook();
        }

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        // 创建一个居中格式
        style.setAlignment(HorizontalAlignment.CENTER);

        //声明列对象
        HSSFCell cell = null;

        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }
        return wb;
    }
    public static List<String[]> getInformation(File xlsFile, int columnNum) throws IOException{
        List<String[]> list = new ArrayList<String[]>();
        if (!xlsFile.getParentFile().exists()) {
            xlsFile.getParentFile().mkdirs();
        }

        //构造 Workbook对象
        Workbook workbook = null;
        try {
            //Excel 2007获取方法
            workbook=new XSSFWorkbook(new FileInputStream(xlsFile));
        } catch(Exception ex){
            //Excel 2003获取方法
            workbook=new HSSFWorkbook(new FileInputStream(xlsFile));
        }
        ///先用InputStream获取excel文件的io流,然后创建一个内存中的excel文件HSSFWorkbook类型对象，这个对象表示了整个excel文件。
        /*HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(xlsFile));*/
        /// 得到Excel工作表对象 Sheet 标识某一页(默认只有sheet1)
        Sheet sheet = workbook.getSheetAt(0);
        //得到总行数
        int count=sheet.getLastRowNum();
        //得到Excel工作表的行
        Row row = sheet.getRow(0);
        //总列数
        //第一行为列名，所以从第二行开始填充的数据
        for(int i=1;i<=count;i++){
            Row row1=sheet.getRow(i);
            String[] line = new String[columnNum];
            Map<String,Object> map = new HashMap<String,Object>();
            for(int j = 0; j< columnNum; j++){
                try{
                    line[j]=getCellFormatValue(row1.getCell(j));
                }catch (Exception e){
                }
            }
            list.add(line);
        }
        return list;
    }

    /**
     * 读取Excel表格表头的内容
     *
     * @param is
     * @return String 表头内容的数组
     */
    public static String[] readExcelTitle(FileInputStream is) {
        try {
            fs = new POIFSFileSystem(is);
            wb = new HSSFWorkbook(fs);
        } catch (IOException e) {
            e.printStackTrace();
        }
        sheet = wb.getSheetAt(0);
        //得到首行的row
        row = sheet.getRow(0);
        // 标题总列数
        int colNum = row.getPhysicalNumberOfCells();
        String[] title = new String[colNum];
        for (int i = 0; i < colNum; i++) {
            title[i] = getCellFormatValue(row.getCell(i));
        }
        return title;
    }

    /**
     * 读取Excel数据内容
     *
     * @param is
     * @return Map 包含单元格数据内容的Map对象
     */
    public static Map<Integer, String> readExcelContent(FileInputStream is) throws IOException {
        Map<Integer, String> content = new HashMap<Integer, String>();
        String str = "";
        try {
            fs = new POIFSFileSystem(is);
            wb = new HSSFWorkbook(is);
        } catch (IOException e) {
            wb= new XSSFWorkbook(is);
            e.printStackTrace();
        }
        sheet = wb.getSheetAt(0);
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        //由于第0行和第一行已经合并了  在这里索引从2开始
        row = sheet.getRow(1);
        int colNum = row.getPhysicalNumberOfCells();
        // 正文内容应该从第二行开始,第一行为表头的标题
        for (int i = 1; i <= rowNum; i++) {
            row = sheet.getRow(i);
            int j = 0;
            while (j < colNum) {
                getCellFormatValue(row.getCell((short) j));
                str += getCellFormatValue(row.getCell((short) j)).trim() + "-";
                j++;
            }
            content.put(i, str);
            str = "";
        }
        return content;
    }

    /**
     * 获取单元格数据内容为字符串类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    private static String getStringCellValue(HSSFCell cell) {
        String strCell = "";
        switch (cell.getCellTypeEnum()) {
            case STRING:
                strCell = cell.getStringCellValue();
                break;
            case NUMERIC:
                strCell = String.valueOf(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                strCell = String.valueOf(cell.getBooleanCellValue());
                break;
            case BLANK:
                strCell = "";
                break;
            default:
                strCell = "";
                break;
        }
        if (strCell.equals("") || strCell == null) {
            return "";
        }
        if (cell == null) {
            return "";
        }
        return strCell;
    }

    /**
     * 获取单元格数据内容为日期类型的数据
     *
     * @param cell Excel单元格
     * @return String 单元格数据内容
     */
    private static String getDateCellValue(HSSFCell cell) {
        String result = "";
        try {
            CellType cellType = cell.getCellTypeEnum();
            if (cellType == NUMERIC ) {
                Date date = cell.getDateCellValue();
                result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1)
                        + "-" + date.getDate();
            } else if (cellType == STRING) {
                String date = getStringCellValue(cell);
                result = date.replaceAll("[年月]", "-").replace("日", "").trim();
            } else if (cellType == BLANK) {
                result = "";
            }
        } catch (Exception e) {
            System.out.println("日期格式不正确!");
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 根据HSSFCell类型设置数据
     *
     * @param cell
     * @return
     */
    private static String getCellFormatValue(Cell cell) {
        String cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellTypeEnum()) {
                // 如果当前Cell的Type为NUMERIC
                case NUMERIC: {

                    NumberFormat numberFormat = NumberFormat.getNumberInstance();
                    numberFormat.setGroupingUsed(false);
                    String format = numberFormat.format(cell.getNumericCellValue());
                    System.out.println(format);
                    cellvalue = format;
                    break;
                }
                case FORMULA:  {
                    // 判断当前的cell是否为Date
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellvalue = sdf.format(date);
                    }
                    // 如果是纯数字
                    else {
                        // 取得当前Cell的数值
                        cellvalue = String.valueOf((int)cell.getNumericCellValue());
                    }
                    break;
                }
                // 如果当前Cell的Type为STRIN
                case STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                // 默认的Cell值
                default:
                    cellvalue = " ";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;

    }
    /**
     * 设置表头
     * @param title
     * @param sheet
     * @param style
     */
    public static void setTitle(String[] title,HSSFSheet sheet,HSSFCellStyle style,int rownum){
        HSSFRow row = sheet.createRow(rownum);
        for(int i=0;i<title.length;i++){
            HSSFCell cell = row.createCell(i);
            if(rownum == 0){
                if("图标Url".equals(title[i]) || "模块路径".equals(title[i])){
                    // 设置列宽
                    sheet.setColumnWidth(i, 12000);
                }else if("id".equals(title[i]) || "父级id".equals(title[i])|| "显示顺序".equals(title[i])){
                    // 设置列宽
                    sheet.setColumnWidth(i, 3000);
                }else{
                    // 设置列宽
                    sheet.setColumnWidth(i, 6000);
                }
            }
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }
    }
    /**
     * 获取workbook
     * @param type
     * @return
     */
    public static HSSFWorkbook getWorkbook(Integer type,String[] title,String sheetName){
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        //先设置为自动换行
        style.setWrapText(true);
        HSSFSheet sheet = hssfWorkbook.createSheet(sheetName);
        ExcelUtil.setTitle(title,sheet,style,0);
        return hssfWorkbook;
    }
}
