import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.titledList.TitledListSheetPage;

public class SimpleWorkbookJava extends AbsWorkbookJavaObj {

    private TitledListSheetPage<SimpleSheet> sheetPage;

    public TitledListSheetPage<SimpleSheet> getSheetPage() {
        return sheetPage;
    }

    public void setSheetPage(TitledListSheetPage<SimpleSheet> sheetPage) {
        this.sheetPage = sheetPage;
    }
}
