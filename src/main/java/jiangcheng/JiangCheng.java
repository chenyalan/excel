package jiangcheng;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * Created by 陈亚兰 on 2018/3/28.
 * 容城和雄县的奖惩信息
 */
public class JiangCheng {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
        String filePath="C:\\Users\\Administrator\\Desktop\\jiangcheng";

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
            String officerId=getCellValue(r.getCell(2));
            String all=getCellValue(r.getCell(1));
            String[] splitStr=all.split("；");
            for(int i=0;i<splitStr.length;i++){
                rowNew=sheetNew.createRow(rowNum++);
                cellNew=rowNew.createCell(0);
                cellNew.setCellValue(officerId);
                cellNew=rowNew.createCell(1);
                cellNew.setCellValue(splitStr[i]);
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
