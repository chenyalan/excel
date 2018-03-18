package sheet;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * created by 陈亚兰 on 18-3-13
 * 处理干部成员信息 --雄安新区-家庭成员信息
 */
public class ReadSheet1 {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
       String filePath="C:\\Users\\Administrator\\Desktop\\雄安新区所有文件";

           getFiles(filePath);

    }

    private static void getFiles(String filePath) throws Exception {
        File root=new File(filePath);
        File[] files=root.listFiles();
        for(File file:files){
            System.out.println(file.getAbsolutePath()+",name:"+file.getName());
            System.out.println("================================="+file.getName()+"=========================================");
            InputStream in=new FileInputStream(file);
            Workbook workbook=getWorkBook(in,file.getName());
            String fileType=file.getName().substring(file.getName().lastIndexOf("."));
            if(excel2003.equals(fileType)){
               workbook=readSheet(workbook,2003);
            }else if(excel2007.equals(fileType)){
               workbook=readSheet(workbook,2007);
            }
            FileOutputStream fo = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\date\\"+file.getName()); // 输出到文件
            workbook.write(fo);
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
        Sheet sheet = wb.getSheetAt(2);//读取第一个sheet页表格内容
        Row row;
        Workbook newWork=new XSSFWorkbook();
        Sheet sheetNew=newWork.createSheet();
        Row rowNew;
        Cell cellNew;
        int rowNum=0;
        for(Row r:wb.getSheetAt(1)){
            rowNew=sheetNew.createRow(rowNum++);
            int n=0;
//            for(Cell c:r) {
            for(int j = 0; j <=r.getLastCellNum(); j++){
                Cell c=r.getCell(j);
                cellNew = rowNew.createCell(j);
                if(c==null){
                    continue;
                }
                if (c.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    if ( j==9 ) {
                            double d = c.getNumericCellValue();
                            Date date = HSSFDateUtil.getJavaDate(d);
                            cellNew.setCellType(Cell.CELL_TYPE_STRING);
                            SimpleDateFormat dformat = new SimpleDateFormat("yyyy-MM");
                            String value = dformat.format(date);
                            cellNew.setCellValue(value);
                    }else{
                        cellNew.setCellValue(getCellValue(c));
                    }
                }else{
                    cellNew.setCellValue(getCellValue(c));
                }
                n++;
            }
        }
        return newWork;
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
            String value;
//            if(HSSFDateUtil.isCellDateFormatted(cell)){
//                double d = cell.getNumericCellValue();
//                        Date date = HSSFDateUtil.getJavaDate(d);
//                        SimpleDateFormat dformat=new SimpleDateFormat("yyyy-MM-dd");
//                        value= dformat.format(date);
//                        return value;
//            }else{
//                value=String.valueOf(cell.getNumericCellValue());
//            }
            value=String.valueOf(cell.getNumericCellValue());
            return value;

        }
        return "";
    }

}
