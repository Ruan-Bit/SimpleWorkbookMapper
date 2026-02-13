import com.simpleWorkbook.annotations.TitleField;
import com.simpleWorkbook.model.AbsSheetJavaObj;

import java.util.List;

public class ComplexSheet extends AbsSheetJavaObj {

    @TitleField("基本信息")
    private ComplexBaseInfo complexBaseInfo;

    @TitleField("薪资实发")
    private List<ComplexSalary> complexSalaries;

    public ComplexBaseInfo getComplexBaseInfo() {
        return complexBaseInfo;
    }

    public void setComplexBaseInfo(ComplexBaseInfo complexBaseInfo) {
        this.complexBaseInfo = complexBaseInfo;
    }

    public List<ComplexSalary> getComplexSalaries() {
        return complexSalaries;
    }

    public void setComplexSalaries(List<ComplexSalary> complexSalaries) {
        this.complexSalaries = complexSalaries;
    }

}
