import java.io.File;

public class MainTest {


    public static void main(String[] args) {
        ClassLoader classLoader = MainTest.class.getClassLoader();

        String path = classLoader.getResource("testExcels/simple.xlsx").getPath();
        File file = new File(path);
        try {

        }catch (Exception e){
            e.printStackTrace();
        }

    }
}
