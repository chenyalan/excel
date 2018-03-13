import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.Messaging.SYNC_WITH_TRANSPORT;

import java.io.*;

/**
 * created by 陈亚兰 on 18-3-13
 * 处理干部成员信息
 */
public class Main {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
       String filePath="/home/cyl/桌面/yfff";

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
            FileOutputStream fo = new FileOutputStream("/home/cyl/桌面/ck/"+file.getName()); // 输出到文件
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
        Sheet sheet = wb.getSheetAt(1);//读取第一个sheet页表格内容
        Object value = null;
        Row row = null;
        Cell cell = null;
        String officerName;
        String officerId;
        String index;
        int lastNum=0;


        //得到这个sheet一共有多少行数据，因为sheet.getLastRow不准
        for(int i=3;;i++){
            index=getExcel2007Value((XSSFCell) sheet.getRow(i).getCell(0));
            if(index.trim().equals("end")){
                lastNum=i;
                break;
            }
        }

        int endRow=0;
        for(int i=3;i<lastNum;i++){
            officerName=getExcel2007Value((XSSFCell) sheet.getRow(i).getCell(1));
            officerId=getExcel2007Value((XSSFCell) sheet.getRow(i).getCell(2));
            for(int j=i+1;j<=lastNum-1;j++){
                String a,b;
                a=b="";
                Cell cella=sheet.getRow(j).getCell(1);
                Cell cellb=sheet.getRow(j).getCell(2);
                a=getExcel2007Value((XSSFCell) cella);
                b=getExcel2007Value((XSSFCell) cellb);
                if((a.trim().equals("")||b.trim().equals(""))&&(b.trim().equals("")||b==null)){
                    cella.setCellValue(officerName);
                    cellb.setCellValue(officerId);
                }else{
                    i=j;
                    break;
                }
            }
            officerName=getExcel2007Value((XSSFCell) sheet.getRow(i).getCell(1));
            officerId=getExcel2007Value((XSSFCell) sheet.getRow(i).getCell(2));
            for(int j=i+1;j<=lastNum-1;j++){
                String a,b;
                a=b="";
                Cell cella=sheet.getRow(j).getCell(1);
                Cell cellb=sheet.getRow(j).getCell(2);
                a=getExcel2007Value((XSSFCell) cella);
                b=getExcel2007Value((XSSFCell) cellb);
                if((a.trim().equals("")||b.trim().equals(""))&&(b.trim().equals("")||b==null)){
                    cella.setCellValue(officerName);
                    cellb.setCellValue(officerId);
                }else{
                    i=j;
                    break;
                }
            }

            if(i==lastNum-1){
                return wb;
            }
            i--;
        }
        return wb;
    }
    //2003
    public static String getExcel2003Value(HSSFCell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING://字符串类型
                cellValue = cell.getStringCellValue();
                if(cellValue.trim().equals("")||cellValue.trim().length()<=0)
                    cellValue=" ";
                break;
            case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_FORMULA: //公式
                cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                cellValue= "";
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                cellValue = Boolean.toString(cell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                cellValue = String.valueOf(cell.getErrorCellValue());
                break;
            default:
                break;
        }
        return cellValue;
    }

    public static String getExcel2007Value(XSSFCell cell) {
        String cellValue = "";
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING://字符串类型
                cellValue = cell.getStringCellValue();
                if(cellValue.trim().equals("")||cellValue.trim().length()<=0)
                    cellValue=" ";
                break;
            case XSSFCell.CELL_TYPE_NUMERIC: //数值类型
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case XSSFCell.CELL_TYPE_FORMULA: //公式
                cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                cellValue= "";
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                cellValue = Boolean.toString(cell.getBooleanCellValue());
                break;
            case XSSFCell.CELL_TYPE_ERROR:
                cellValue = String.valueOf(cell.getErrorCellValue());
                break;
            default:
                break;
        }
        return cellValue;
    }

}
