package rongchengcunganbu;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * Created by 陈亚兰 on 2018/3/15.
 * 判断八于乡
 */
public class CreateWorkbook {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
        String filePath="C:\\Users\\Administrator\\Desktop\\a";

        getFiles(filePath);

    }

    private static void getFiles(String filePath) throws Exception {
        File root=new File(filePath);
        File[] files=root.listFiles();
        for(File file:files){
//            System.out.println(file.getAbsolutePath()+",name:"+file.getName());
            System.out.println("================================="+file.getName()+"=========================================");
            InputStream in=new FileInputStream(file);
            Workbook workbook=getWorkBook(in,file.getName());
            String fileType=file.getName().substring(file.getName().lastIndexOf("."));
            if(excel2003.equals(fileType)){
                workbook=readSheet(workbook,2003);
            }else if(excel2007.equals(fileType)){
                workbook=readSheet(workbook,2007);
            }
            Sheet sheet=workbook.getSheetAt(0);

            Cell cell;
            Cell cellNext;

            Row row;
            StringBuffer sb=new StringBuffer();
            Map<String,String> map=new LinkedHashMap<String, String>();
            for(int i=3;i<=8;i++){
                row=sheet.getRow(i);
                for(int j=0;j<row.getLastCellNum();j++){
                    cell=row.getCell(j);
                    switch (j){
                        case 0:
                        case 3:
                        case 7:
                            cellNext=row.getCell(j+1);
                            map.put(getCellValue(cell),getCellValue(cellNext));
                    }
                }
            }
            List<String> args= Arrays.asList("关系","姓名","身份证号","性别","民族","政治面貌","出生日期","工作单位","职务","职级");
            ArrayList<List<String>> arg=new ArrayList<List<String>>();
            //保存数据
            Workbook wbNew=new XSSFWorkbook();
            Sheet sheetNew=wbNew.createSheet();
            Row rowNew=sheetNew.createRow(0);

            Cell cellNew=rowNew.createCell(0);
            cellNew.setCellValue(file.getName());
            rowNew=sheetNew.createRow(1);
            for(Map.Entry<String,String> ma:map.entrySet()){
                System.out.print(ma.getKey()+":"+ma.getValue()+"\t\t");
            }
             for(int i=11;i<sheet.getLastRowNum();i++){
                 List<String> list=new ArrayList<String>();
                 row=sheet.getRow(i);
                 for(int j=0;j<10;j++){
                     cell=row.getCell(j);
                     list.add(getCellValue(cell));
                     if(cell==null) continue;
                 }
                 arg.add(list);
             }
             System.out.println("------------家庭成员--------------");
             for(int i=0;i<arg.size();i++){
                 List<String> list=arg.get(i);
                 for(int j=0;j<list.size();j++){
                     System.out.print(list.get(j)+"\t\t");
                 }
                 System.out.print("\n");
             }
//            for(int i=0;i<row.getLastCellNum();i++){
//                Cell c=row.getCell(i);
//                cellNew=rowNew.createCell(i);
//                if(c==null)continue;
//                cellNew.setCellValue(getCellValue(c));
//            }
//            FileOutputStream fo = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\congcheng\\"+file.getName()); // 输出到文件
//            wbNew.write(fo);
        }

    }

    private static Workbook getWorkBook(InputStream in, String name) throws IOException {
        String fileType=name.substring(name.lastIndexOf("."));
        Workbook workbook=null;
        if(excel2003.equals(fileType)){
            workbook=new HSSFWorkbook(in);
        }else if(excel2007.equals(fileType)){
            workbook=new XSSFWorkbook(in);
        }
        return workbook;
    }


    //读取excel文件
    private static Workbook readSheet(Workbook wb,int type) throws FileNotFoundException {
        Sheet sheet = wb.getSheetAt(0);//读取第一个sheet页表格内容
        Object value = null;
        Row row = null;
        Cell cell = null;
        String officerName;
        String officerId;
        String index;
        int lastNum=0;
        row=sheet.getRow(10);
        StringBuffer sb=new StringBuffer();
        for(Cell c:row){
            sb.append(getCellValue(c)+"  ");
        }
        System.out.println(sb.toString());

        return wb;
    }

    public static String getCellValue(Cell cell){

        if(cell == null) return "";

        if(cell.getCellType() == Cell.CELL_TYPE_STRING){

            return cell.getStringCellValue();

        }else if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN){

            return String.valueOf(cell.getBooleanCellValue());

        }else if(cell.getCellType() == Cell.CELL_TYPE_FORMULA){

            return cell.getCellFormula() ;

        }else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }
}
