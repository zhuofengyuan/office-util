package com.fengtoos.customer.officeutil.util;

import com.fengtoos.customer.officeutil.resp.Result;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 读取上篇中的xls文件的内容，并打印出来
 *
 * @author Administrator
 */
public class ExcelUtil {

//    public static void main(String[] args) throws Exception {
//        Result rs = ExcelUtil.readTable(new File("C:\\Users\\feng\\Desktop\\data.xlsx"), 6);
//        ExcelUtil.createDocument(rs.getData(), "F:\\template\\模板.xlsx");
//        String str = "1、3号点位于史家庄村、封家坝村、马蹄湾社区交界处，为本权属交界线的起点，向南方向738.58米到达13号点。";
//        String str1 = "10、69号点位于马蹄湾社区、嘉陵江、十天高速公路交界处，为本权属交界线的终点。";
//        System.out.println(str.split("号点").length);
//        System.out.println(Arrays.toString(str.split("号点")));
//        System.out.println(str.substring(str.indexOf("到达") + 2, str.lastIndexOf("号点")));
//        System.out.println(str1.substring(0, str1.indexOf("、")));
//        System.out.println(str1.substring(str1.indexOf("、") + 1, str1.indexOf("号点")));
//        System.out.println(str1.substring(str1.indexOf("点") + 1));
//        String test = str1;
//        if(test.indexOf("终点") == -1){
//            if(str.split("号点").length > 2){
//                System.out.println(test.substring(0, test.indexOf("、"))); //序号
//                System.out.println(test.substring(test.indexOf("、") + 1, test.indexOf("号点"))); //起点
//                System.out.println(test.substring(test.indexOf("点") + 1, test.indexOf("到达")));//正文
//                System.out.println(test.substring(test.indexOf("到达") + 2, test.lastIndexOf("号点")));
//            }
//        } else {
//            System.out.println(test.substring(0, test.indexOf("、"))); //序号
//            System.out.println(test.substring(test.indexOf("、") + 1, test.indexOf("号点"))); //起点
//            System.out.println(test.substring(test.indexOf("点") + 1, test.indexOf("。")));
//            System.out.println("终点");
//        }
//    }

    //通过对单元格遍历的形式来获取信息 ，这里要判断单元格的类型才可以取出值
    public static Result readTable(File file, Integer index) throws Exception {
        InputStream ips = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(ips);
        XSSFSheet sheet = wb.getSheetAt(0);
        //第一個表的數據
        List<ArrayList<Object>> list = getSheetData(sheet);

        List<ArrayList<Object>> rs = new ArrayList<ArrayList<Object>>();
        String msg = "拆分成功！！";
        boolean isSuccess = true;

        //每一行
        int z = 0, xlen = 1;
        for (ArrayList<Object> item : list) {
            //每一个单元格
            for (int i = 0; i < item.size(); i++) {
                //找到需要分割的单元格
                if (i + 1 == index) {
                    String value = item.get(i).toString();
                    String[] split = value.split("\\n");
                    for (String str : split) {
                        ArrayList<Object> obj = new ArrayList<>(item);
                        obj.set(i, str);
                        ArrayList<Object> newo = new ArrayList<>();
                        for(int k = 0, klen = obj.size(); k < klen; k++){
                            //分割前一个单元格
                            if(k + 1 == index){
                                String s = obj.get(k).toString();
                                try{
                                    //标题
                                    if(z == 0){
                                        newo.add("序号");
                                        newo.add("起点号");
                                        newo.add("终点号");
                                        newo.add(s);
                                    } else {
                                        if(s.split("号点").length == 2 && s.indexOf("终点") == -1){
                                            throw new StringIndexOutOfBoundsException();
                                        }

                                        if(s.indexOf("终点") == -1){
                                            if(s.split("号点").length > 2){
                                                newo.add(s.substring(0, s.indexOf("、")));
                                                newo.add(s.substring(s.indexOf("、") + 1, s.indexOf("号点")));
                                                newo.add(s.substring(s.indexOf("到达") + 2, s.lastIndexOf("号点")));
                                                newo.add(s.substring(s.indexOf("点") + 1, s.indexOf("到达")));
                                            }
                                        } else {
                                            if(s.split("号点").length == 2){
                                                newo.add(s.substring(0, s.indexOf("、")));
                                                newo.add(s.substring(s.indexOf("、") + 1, s.indexOf("号点")));
                                                newo.add("");
                                                newo.add(s.substring(s.indexOf("点") + 1, s.indexOf("。")));
                                            }
                                        }
                                    }
                                } catch (StringIndexOutOfBoundsException e){
                                    msg = "第" + xlen +  "行\n宗地号：" + newo.get(0) + "\n数据：" + s + "\n不符合<、>,<号点>,<。中文句号>转换要求！！！";
                                    System.out.println(msg);
                                    isSuccess = false;
                                    break;
                                }
                            } else {
                                newo.add(obj.get(k));
                            }
                        }
                        rs.add(newo);
                    }
                }
            }
            z++;
            xlen++;
        }
        wb.close();
        return Result.normal(msg, isSuccess, rs);
    }

    public static void createDocument(List<ArrayList<Object>> list, String outPath) throws IOException {
        //创建工作薄对象
        XSSFWorkbook workbook = new XSSFWorkbook();//这里也可以设置sheet的Name
        //创建工作表对象
        XSSFSheet sheet = workbook.createSheet();
        //创建工作表的行
        Map<Integer, Integer> maxWidth = new HashMap<Integer, Integer>();
        for (int i = 0, rlen = list.size(); i < rlen; i++) {
            XSSFRow row = sheet.createRow(i);
            ArrayList<Object> cells = list.get(i);
            for (int j = 0, jlen = cells.size(); j < jlen; j++) {
                XSSFCell cell = row.createCell(j);
                String value = (cells.get(j) + "").trim();
                if (i == 0) {
                    maxWidth.put(j, value.getBytes().length * 256 + 200);
                } else {
                    int x = value.getBytes().length * 256 + 200;
                    if (x > 15000) {
                        x = 15000;
                    }
                    maxWidth.put(j, Math.max(x, maxWidth.get(j)));
                }
                cell.setCellValue(value);
                cell.setCellStyle(setStyle(workbook));
            }
        }

        // 列宽自适应
        for (int i = 0; i < maxWidth.keySet().size(); i++) {
            sheet.setColumnWidth(i, maxWidth.get(i));
        }

        workbook.setSheetName(0, "sheet1");//设置sheet的Name

        //文档输出
        FileOutputStream out = new FileOutputStream(outPath);
        workbook.write(out);
        out.close();
    }

    private static List<ArrayList<Object>> getSheetData(XSSFSheet sheet) {
        List<ArrayList<Object>> list = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
        for (Iterator ite = sheet.rowIterator(); ite.hasNext(); ) {
            XSSFRow row = (XSSFRow) ite.next();
            ArrayList<Object> rowm = new ArrayList<>();
            for (Iterator cellite = row.cellIterator(); cellite.hasNext(); ) {
                XSSFCell cell = (XSSFCell) cellite.next();
                if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                    rowm.add(sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue())));
                } else {
                    rowm.add(cell);
                }
            }
            list.add(rowm);
        }
        return list;
    }

    private static XSSFCellStyle setStyle(XSSFWorkbook workbook) {
        XSSFCellStyle style = workbook.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);//上边框
        style.setBorderBottom(BorderStyle.THIN);//下边框
        style.setBorderLeft(BorderStyle.THIN);//左边框
        style.setBorderRight(BorderStyle.THIN);//右边框
        style.setTopBorderColor(new XSSFColor(java.awt.Color.black));//上边框颜色
        style.setBottomBorderColor(new XSSFColor(java.awt.Color.black));//下边框颜色
        style.setLeftBorderColor(new XSSFColor(java.awt.Color.black));//左边框颜色
        style.setRightBorderColor(new XSSFColor(java.awt.Color.black));//右边框颜色
        style.setWrapText(true);//自动换行
        style.setAlignment(HorizontalAlignment.CENTER); // 水平居中
        style.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
        return style;
    }
}