import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.simpleWorkbook.SimpleWorkbookMapper;
import com.simpleWorkbook.model.titledList.TitledListSheetPageObj;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

public class ComplexWriteTest {

    private static final Random RANDOM = new Random();
    private static final String[] NAMES = {"张三", "李四", "王五", "赵六", "孙七", "周八", "吴九", "郑十"};
    private static final String[] SEXES = {"男", "女"};
    private static final String[] DEPARTMENTS = {"技术部", "销售部", "人事部", "财务部", "市场部"};

    public static void main(String[] args) {
        try {
            // 1. 生成模拟数据
            ComplexWorkbook complexWorkbook = generateMockData();

            Gson gson = new Gson();
            System.out.println(gson.toJson(complexWorkbook));

            // 2. 写入 Excel 文件
            File inputFile = new File(Paths.get("").toAbsolutePath().toString() + "/complexOutPutSheet.xlsx");
            if (inputFile.exists()) {
                inputFile.delete();
            }
            inputFile.createNewFile();
            try (Workbook workbook = SimpleWorkbookMapper.writeWorkbook(complexWorkbook);
                 OutputStream outputStream = Files.newOutputStream(inputFile.toPath())
            ) {
                workbook.write(outputStream);
            }
            System.out.println("\nExcel 文件已生成: " + inputFile.toPath().toString());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成模拟数据
     */
    public static ComplexWorkbook generateMockData() {
        ComplexWorkbook workbook = new ComplexWorkbook();

        TitledListSheetPageObj<ComplexSheet> complexSheetTitledListSheetPageObj = new TitledListSheetPageObj<>();
        List<ComplexSheet> complexSheetList = new ArrayList<>();
        complexSheetTitledListSheetPageObj.setData(complexSheetList);
        workbook.setComplexSheet(complexSheetTitledListSheetPageObj);

        // 1
        ComplexSheet complexSheet1 = new ComplexSheet();
        complexSheet1.setComplexBaseInfo(generateBaseInfo(1));

        List<ComplexSalary> salaryList = new ArrayList<>();
        for (int j = 0; j < 3; j++) {
            salaryList.add(generateSalary(j));
        }

        complexSheet1.setComplexSalaries(salaryList);
        complexSheetList.add(complexSheet1 );

        // 2
        ComplexSheet complexSheet2 = new ComplexSheet();
        complexSheet2.setComplexBaseInfo(generateBaseInfo(2));

        List<ComplexSalary> salaryList2 = new ArrayList<>();
        for (int j = 0; j < 2; j++) {
            salaryList2.add(generateSalary(j));
        }

        complexSheet2.setComplexSalaries(salaryList2);
        complexSheetList.add(complexSheet2);

        // 3
        ComplexSheet complexSheet3 = new ComplexSheet();
        complexSheet3.setComplexBaseInfo(generateBaseInfo(3));

        List<ComplexSalary> salaryList3 = new ArrayList<>();
        for (int j = 0; j < 4; j++) {
            salaryList3.add(generateSalary(j));
        }

        complexSheet3.setComplexSalaries(salaryList3);
        complexSheetList.add(complexSheet3);

        return workbook;
    }

    /**
     * 生成基础信息
     */
    private static ComplexBaseInfo generateBaseInfo(int index) {
        ComplexBaseInfo info = new ComplexBaseInfo();
        info.setJobNumber("EMP" + String.format("%04d", index + 1));
        info.setName(NAMES[RANDOM.nextInt(NAMES.length)]);
        info.setSex(SEXES[RANDOM.nextInt(SEXES.length)]);
        info.setDepartment(DEPARTMENTS[RANDOM.nextInt(DEPARTMENTS.length)]);
        return info;
    }

    /**
     * 生成薪资信息
     */
    private static ComplexSalary generateSalary(int index) {
        ComplexSalary salary = new ComplexSalary();
        int base = 5000 + RANDOM.nextInt(10000);
        int performance = 1000 + RANDOM.nextInt(5000);
        int allowance = 500 + RANDOM.nextInt(2000);
        int bonus = RANDOM.nextInt(5000);
        int total = base + performance + allowance + bonus;

        salary.setBase(String.valueOf(base));
        salary.setPerformance(String.valueOf(performance));
        salary.setAllowance(String.valueOf(allowance));
        salary.setBonus(String.valueOf(bonus));
        salary.setTotal(String.valueOf(total));
        salary.setDate("2024-" + String.format("%02d", RANDOM.nextInt(12) + 1));
        return salary;
    }
}