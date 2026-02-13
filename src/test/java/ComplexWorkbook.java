import com.simpleWorkbook.annotations.SheetField;
import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.titledList.TitledListSheetPageObj;

public class ComplexWorkbook extends AbsWorkbookJavaObj {

    @SheetField("薪资")
    private TitledListSheetPageObj<ComplexSheet> complexSheet;

    public TitledListSheetPageObj<ComplexSheet> getComplexSheet() {
        return complexSheet;
    }

    public void setComplexSheet(TitledListSheetPageObj<ComplexSheet> complexSheet) {
        this.complexSheet = complexSheet;
    }
}
