import com.simpleWorkbook.annotations.TitleField;
import com.simpleWorkbook.model.AbsSheetJavaObj;

public class SimpleSheet extends AbsSheetJavaObj {

    @TitleField(value = "姓名", colWidth = 20)
    private String name;

    @TitleField(value = "年龄", colWidth = 10)
    private String age;

    @TitleField(value = "性别", colWidth = 10)
    private String sex;



    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAge() {
        return age;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }
}
