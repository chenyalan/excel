package judgeziduan;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * Created by 陈亚兰 on 2018/3/15.
 * 判断八于乡
 */
public class JudgeZiduanBayuxiang {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
       // String filePath="C:\\Users\\Administrator\\Desktop\\村干部所有信息存放在一起";
String filePath="C:\\Users\\Administrator\\Desktop\\special";
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
            FileOutputStream fo = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\ddg\\"+file.getName()); // 输出到文件
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
        Sheet sheet = wb.getSheetAt(0);//读取第一个sheet页表格内容
        if(sheet.getRow(1).getCell(0)==null&&sheet.getRow(0).getCell(0)==null) return wb;
        Object value = null;
        Row row = null;
        Cell cell = null;
        String officerName;
        String officerId;
        String index;
        int lastNum=0;
        row=sheet.getRow(9);
        StringBuffer sb=new StringBuffer();
        for(Cell c:row){
            if(c==null)
                continue;
            sb.append(getCellValue(c)+"  ");
        }
        System.out.println(sb.toString());
        wb.setSheetName(0, "name");
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
