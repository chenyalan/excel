package gzll;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by 陈亚兰 on 2018/3/28.
 * 三个县的工作履历 操作第二张Sheet然后输出Sheet3
 */
public class ThreeGzll2 {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
        String filePath="C:\\Users\\Administrator\\Desktop\\threegzll";

        getFiles(filePath);

    }
    private static void getFiles(String filePath) throws Exception {
        File root=new File(filePath);
        File[] files=root.listFiles();
        for(File file:files){
            InputStream in=new FileInputStream(file);
            Workbook workbook=getWorkBook(in,file.getName());
            String fileType=file.getName().substring(file.getName().lastIndexOf("."));
            if(excel2003.equals(fileType)){
                workbook=readSheet(workbook,2003);
            }else if(excel2007.equals(fileType)){
                workbook=readSheet(workbook,2007);
            }
            FileOutputStream fo = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\xd\\"+file.getName()); // 输出到文件
            workbook.write(fo);
        }

    }

    //读取excel文件
    private static Workbook readSheet(Workbook wb,int type) throws FileNotFoundException {
        Sheet sheet = wb.getSheetAt(1);//读取第一个sheet页表格内容
        Object value = null;
        Sheet sheetNew=wb.createSheet();
        Row rowNew;
        Cell cellNew;
        Row row = null;
        int rowNum=0;
        String reg="(19[0-9][0-9]|20[0-1][0-9]).([0-1][0-9])-";
        int total=sheet.getLastRowNum();
        for(int i=0;i<sheet.getLastRowNum();i++){
            row=sheet.getRow(i);
            String all=getCellValue(row.getCell(2));
            String officerId=getCellValue(row.getCell(0));
            String year=getCellValue(row.getCell(1));
            Pattern p=Pattern.compile(reg);
            Matcher m=p.matcher(all);
            m.find();
            int start,end;
            try{
                start=m.start();
                end=m.end();
                rowNew=sheetNew.createRow(rowNum++);
                cellNew=rowNew.createCell(0);
                cellNew.setCellValue(officerId);
                cellNew=rowNew.createCell(1);
                cellNew.setCellValue(year);
                cellNew=rowNew.createCell(2);
                cellNew.setCellValue(all.substring(0,start));
                rowNew=sheetNew.createRow(rowNum++);
                cellNew=rowNew.createCell(0);
                cellNew.setCellValue(officerId);
                cellNew=rowNew.createCell(1);
                cellNew.setCellValue(all.substring(start,end));
                cellNew=rowNew.createCell(2);
                cellNew.setCellValue(all.substring(end,all.length()));
            }catch (Exception e){
                rowNew=sheetNew.createRow(rowNum);
                cellNew=rowNew.createCell(0);
                cellNew.setCellValue(officerId);
                cellNew=rowNew.createCell(1);
                cellNew.setCellValue(year);
                cellNew=rowNew.createCell(2);
                cellNew.setCellValue(all);
                rowNum++;
            }
        }

        return wb;
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
