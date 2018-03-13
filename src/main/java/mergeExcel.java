import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.Messaging.SYNC_WITH_TRANSPORT;

import java.io.*;

/**
 * created by 陈亚兰 on 18-3-13
 * 处理干部成员信息
 */
public class mergeExcel {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
       String filePath="/home/cyl/桌面/lv";

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
            FileOutputStream fo = new FileOutputStream("/home/cyl/桌面/lv2/"+file.getName()); // 输出到文件
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
            Cell c=sheet.getRow(i).getCell(0);
            boolean isMerge=isMergedRegion(sheet,i,c.getColumnIndex());
            if(isMerge){
                index=getMergedRegionValue(sheet,sheet.getRow(i).getRowNum(),c.getColumnIndex());
            }else{
                index=getCellValue(sheet.getRow(i).getCell(0));
            }
            if(index.trim().equals("end")){
                lastNum=i;
                break;
            }
        }
        StringBuffer sb=new StringBuffer();
        Sheet sheetNew=wb.createSheet();
        Row rowNew=null;
        Cell cellNew=cell;
        int k=0;
        for(int i=2;i<=lastNum;i++){
            rowNew=sheetNew.createRow(k++);
            int p=0;
            String val="";
            for(Cell c:sheet.getRow(i)){
                cellNew=rowNew.createCell(p++);
                boolean isMerge=isMergedRegion(sheet,i,c.getColumnIndex());
                if(isMerge){
                    val=getMergedRegionValue(sheet,sheet.getRow(i).getRowNum(),c.getColumnIndex());
                    cellNew.setCellValue(val);
                    sb.append(val+"  ");
                }else{
                    sb.append(getCellValue(c)+"  ");
                    cellNew.setCellValue(getCellValue(c));
                }
            }
            sb.append("\n");
        }
        System.out.println(sb.toString());
        return wb;
    }

    private static String getMergedRegionValue(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();


        for(int i = 0 ; i < sheetMergeCount ; i++){
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();

            if(row >= firstRow && row <= lastRow){

                if(column >= firstColumn && column <= lastColumn){
                    Row fRow = sheet.getRow(firstRow);
                    Cell fCell = fRow.getCell(firstColumn);
                    return getCellValue(fCell) ;
                }
            }
        }

        return null ;
    }

    private static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
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
