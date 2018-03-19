import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by 陈亚兰 on 2018/3/18.
 */
public class DiGuiAllFiles {
    private static void test(String fileDir) {
        List<File> fileList = new ArrayList<File>();
        File file = new File(fileDir);
        File[] files = file.listFiles();// 获取目录下的所有文件或文件夹
        if (files == null) {// 如果目录为空，直接退出
            return;
        }
        // 遍历，目录下的所有文件
        for (File f : files) {
            if (f.isFile()) {
                if(f.getName().endsWith(".xlsx")){
                    fileList.add(f);
                }
//                fileList.add(f);
            } else if (f.isDirectory()) {
                System.out.println(f.getAbsolutePath());
                test(f.getAbsolutePath());
            }
        }
        for (File f1 : fileList) {
            System.out.println(f1.getName());
        }
    }

    public static void main(String[] args) {
        test("C:\\Users\\Administrator\\Desktop\\XA数据314\\XA数据（0310汇总版）\\容城\\村干部\\个人信息\\容城县各乡镇上报书记、村主任信息采集表");
    }
}
