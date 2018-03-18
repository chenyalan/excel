import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by 陈亚兰 on 2018/3/14.
 * 安新----村干部 ---基础信息
 */
public class AxOfficerBasic {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
        String filePath="C:\\Users\\Administrator\\Desktop\\ppp";

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
            FileOutputStream fo = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\aaa\\"+file.getName()); // 输出到文件
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
        Object value = null;
        Row row = null;
        Cell cell = null;
        String officerName;
        String officerId;
        String index;
        int lastNum=0;
        Sheet sheetNew=wb.createSheet();
        Row rowNew;
        Cell cellNew;
        StringBuffer sb=new StringBuffer();
        row=sheet.getRow(1);
//        for(Cell c:row){
//          sb.append(getCellValue(c)+"   ");
//        }
        int j=0;
        int n=1;
        int k;
        for(int i=1;i<=sheet.getLastRowNum();i++){
            k=0;
            rowNew=sheetNew.createRow(i-1);
            for(Cell c:sheet.getRow(i)){
                if(i==1){
                    cellNew=rowNew.createCell(k);
                    switch (k){
                        case 0:
                            cellNew.setCellValue("index");break;
                        case 1:
                            cellNew.setCellValue("name");break;
                        case 2:
                            cellNew.setCellValue("id_card");break;
                        case 3:
                            cellNew.setCellValue("sex");break;
                        case 4:
                            cellNew.setCellValue("nation");break;
                        case 5:
                            cellNew.setCellValue("ji_guan");break;
                        case 6:
                            cellNew.setCellValue("birth");break;
                        case 7:
                            cellNew.setCellValue("age");break;
                        case 8:
                            cellNew.setCellValue("edu");break;
                        case 9:
                            cellNew.setCellValue("healthy");break;
                        case 10:
                            cellNew.setCellValue("marriage");break;
                        case 11:
                            cellNew.setCellValue("telephone");break;
                        case 12:
                            cellNew.setCellValue("address");break;
                        case 13:
                            cellNew.setCellValue("polity");break;
                        case 14:
                            cellNew.setCellValue("join_party");break;
                        case 15:
                            cellNew.setCellValue("join_work");break;
                        case 16:
                            cellNew.setCellValue("company");break;
                        case 17:
                            cellNew.setCellValue("duty");break;
                        default:
                            cellNew.setCellValue("duty_level");
                    }
                }else{
                    cellNew=rowNew.createCell(k);
                    cellNew.setCellValue(getCellValue(c));
                }
                k++;
            }
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
//            if(DateUtil.isCellDateFormatted(cell)){
//                DataFormatter dataFormatter=new DataFormatter();
//                Format format=dataFormatter.createFormat(cell);
//                Date date=cell.getDateCellValue();
//                String str=format.format(date);
//                return str;
//            }
            short format=cell.getCellStyle().getDataFormat();
            SimpleDateFormat sdf=null;
            if(format==57){
                sdf=new SimpleDateFormat("yyyy年MM月");
                double value=cell.getNumericCellValue();
                Date date=DateUtil.getJavaDate(value);
                return sdf.format(date);
            }

            return String.valueOf(cell.getNumericCellValue());

        }else if(cell.getCellType()==Cell.CELL_TYPE_STRING){
            return   cell.getRichStringCellValue().getString();
        }
        return "";
    }
}
