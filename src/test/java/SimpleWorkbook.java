import com.simpleWorkbook.annotations.SheetField;
import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.titledList.TitledListSheetPageObj;

public class SimpleWorkbook extends AbsWorkbookJavaObj {

    @SheetField("Sheet1")
    private TitledListSheetPageObj<SimpleSheet> sheetPage;

    public TitledListSheetPageObj<SimpleSheet> getSheetPage() {
        return sheetPage;
    }

    public void setSheetPage(TitledListSheetPageObj<SimpleSheet> sheetPage) {
        this.sheetPage = sheetPage;
    }
}
