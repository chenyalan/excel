package xueli;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created by 陈亚兰 on 2018/3/28.
 *  三县学历信息
 */
public class XueLi {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
        String filePath="C:\\Users\\Administrator\\Desktop\\threexueli";

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
            FileOutputStream fo = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\dt\\"+file.getName()); // 输出到文件
            workbook.write(fo);
        }

    }

    //读取excel文件
    private static Workbook readSheet(Workbook wb,int type) throws FileNotFoundException {
        Sheet sheet = wb.getSheetAt(0);//读取第一个sheet页表格内容
        Object value = null;
        Sheet sheetNew=wb.createSheet();
        Row rowNew;
        Cell cellNew;
        Row row = null;
        Cell cell = null;
        row=sheet.getRow(2);
        StringBuffer sb=new StringBuffer();
        int rowNum=0;
        for(Row r:sheet){
          String idCard=getCellValue(r.getCell(0));
          String qrzIndex=getCellValue(r.getCell(5));//全日制枚举索引
          String zzIndex=getCellValue(r.getCell(6));//在职枚举索引
            String qrzSchool=getCellValue(r.getCell(2));//全日制学校
            String zzSchool=getCellValue(r.getCell(4));//在职学校
            //FullDay 全日制 0  OnWork 在职 1
            List<String> qrz=Arrays.asList(idCard,"0",qrzIndex,qrzSchool);
            List<String> zz= Arrays.asList(idCard,"1",zzIndex,zzSchool);
            for(int i=0;i<2;i++){
                rowNew=sheetNew.createRow(rowNum++);
                switch (i){
                    case 0:
                        for(int j=0;j<qrz.size();j++){
                            cellNew=rowNew.createCell(j);
                            cellNew.setCellValue(qrz.get(j));
                        }
                        break;
                    case 1:
                        for(int j=0;j<zz.size();j++){
                            cellNew=rowNew.createCell(j);
                            cellNew.setCellValue(zz.get(j));
                        }
                }
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
