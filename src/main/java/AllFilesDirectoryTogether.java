import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * created by 陈亚兰 on 18-3-13
 * 雄安新区-个人信息下所有子目录所有.xlsx文件，存放在58目录下
 */
public class AllFilesDirectoryTogether {
    private final static String excel2003=".xls";
    private final static String excel2007=".xlsx";
    public static void main(String[] args) throws Exception {
       String filePath="C:\\Users\\Administrator\\Desktop\\xx";
       String toPath="C:\\Users\\Administrator\\Desktop\\netw\\";
//           getFiles(filePath);
        test(filePath,toPath);
    }

    private static void test(String fileDir,String toPath) throws IOException {
        List<File> fileList = new ArrayList<File>();
        File fi= new File(fileDir);
        File[] files = fi.listFiles();// 获取目录下的所有文件或文件夹
        if (files == null) {// 如果目录为空，直接退出
            return;
        }
        // 遍历，目录下的所有文件
        for (File f : files) {
            if (f.isFile()) {
                if(f.getName().endsWith(".xlsx")||f.getName().toLowerCase().endsWith(".xls")){
                    fileList.add(f);
                }
            } else if (f.isDirectory()) {
                System.out.println(f.getAbsolutePath());
                test(f.getAbsolutePath(),toPath);
            }
        }
        for (File file : fileList) {
            System.out.println(file.getAbsolutePath()+",name:"+file.getName());
            System.out.println("================================="+file.getName()+"=========================================");
            InputStream in=new FileInputStream(file);
            Workbook workbook=getWorkBook(in,file.getName());
            FileOutputStream fo = new FileOutputStream(toPath+file.getName()); // 输出到文件
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


        //得到这个sheet一共有多少行数据，因为sheet.getLastRow不准
//        for(int i=3;;i++){
//            index=getCellValue(sheet.getRow(i).getCell(0));
//            if(index.trim().equals("end")){
//                lastNum=i;
//                break;
//            }
//        }
        lastNum=2;  //很奇怪还就是2，sheet.getlastRow都不行
        int endRow=0;
        for(int i=1;i<=lastNum;i++){
            officerName=getCellValue(sheet.getRow(i).getCell(1));
            officerId=getCellValue(sheet.getRow(i).getCell(2));
            for(int j=i+1;j<=lastNum-1;j++){
                String a,b;
                a=b="";
                Cell cella=sheet.getRow(j).getCell(1);
                Cell cellb=sheet.getRow(j).getCell(2);
                a=getCellValue(cella);
                b=getCellValue(cellb);
                if((a.trim().equals("")||b.trim().equals(""))&&(b.trim().equals("")||b==null)){
                    cella.setCellValue(officerName);
                    cellb.setCellValue(officerId);
                }else{
                    i=j;
                    break;
                }
            }
            officerName=getCellValue(sheet.getRow(i).getCell(1));
            officerId=getCellValue( sheet.getRow(i).getCell(2));
            for(int j=i+1;j<=lastNum-1;j++){
                String a,b;
                a=b="";
                Cell cella=sheet.getRow(j).getCell(1);
                Cell cellb=sheet.getRow(j).getCell(2);
                a=getCellValue(cella);
                b=getCellValue(cellb);
                if((a.trim().equals("")||b.trim().equals(""))&&(b.trim().equals("")||b==null)){

                    try{
                        cella.setCellValue(officerName);
                    }catch (Exception e){
                        System.out.println(officerName);
                    }
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
