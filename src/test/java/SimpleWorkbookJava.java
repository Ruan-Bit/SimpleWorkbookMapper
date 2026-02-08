import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.titledList.TitledListAbsSheetPageObj;

public class SimpleWorkbookJava extends AbsWorkbookJavaObj {

    private TitledListAbsSheetPageObj<SimpleSheet> sheetPage;

    public TitledListAbsSheetPageObj<SimpleSheet> getSheetPage() {
        return sheetPage;
    }

    public void setSheetPage(TitledListAbsSheetPageObj<SimpleSheet> sheetPage) {
        this.sheetPage = sheetPage;
    }
}
