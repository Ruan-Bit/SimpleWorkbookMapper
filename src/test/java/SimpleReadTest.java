
import com.google.gson.Gson;
import com.simpleWorkbook.SimpleWorkbookMapper;

import java.io.File;

public class SimpleReadTest {

    public static void main(String[] args) {
        // 1. 读取 Excel 文件
        File inputFile = new File(SimpleWorkbookMapper.class.getClassLoader().getResource("testExcels/simpleSheet.xlsx").getPath());

        try {
            // 2. 读取 Excel 到 Java 对象
            SimpleWorkbook workbookJava = SimpleWorkbookMapper.readWorkbook(SimpleWorkbook.class, inputFile);

            // 3. 打印读取的数据
            Gson gson = new Gson();
            System.out.println(gson.toJson(workbookJava));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
