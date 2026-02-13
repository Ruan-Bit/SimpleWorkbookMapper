import com.google.gson.Gson;
import com.simpleWorkbook.SimpleWorkbookMapper;

import java.io.File;

public class ComplexReadTest {


    public static void main(String[] args) {
        // 1. 读取 Excel 文件
        File inputFile = new File(SimpleWorkbookMapper.class.getClassLoader().getResource("testExcels/complexSheet.xlsx").getPath());

        try {
            // 2. 读取 Excel 到 Java 对象
            ComplexWorkbook workbookJava = SimpleWorkbookMapper.readWorkbook(ComplexWorkbook.class, inputFile);

            // 3. 打印读取的数据
            System.out.println(new Gson().toJson(workbookJava));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
