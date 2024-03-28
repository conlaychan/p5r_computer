package persona;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

public class Persona4GoldenComputer {
    public static void computer() {
        System.out.println("Persona 4 Golden：");

        Context context = new Context("Persona4Golden_material.xlsx", "Persona4Golden_computed.xlsx");
        File file = new File(context.resultFile);
        if (file.exists() && !file.delete()) {
            System.err.println("请先删除当前目录下的" + context.resultFile + "文件！");
            return;
        }

        System.out.println("正在读取和分析素材文件：" + context.materialFile);
        readPersonas(context);

        System.out.println("正在计算合成结果。。。");
        Common.compute(context);

        System.out.println("正在保存计算结果：" + context.resultFile);
        Common.saveExcel(context);

        System.out.println("完成！");
    }

    private static void readPersonas(Context context) {
        try (XSSFWorkbook wb = new XSSFWorkbook(context.materialFile)) {
            // 素材面具
            Common.readMaterialPersonas(context, wb.getSheetAt(0));
            // 塔罗牌二合一
            Common.readArcanaRules(context, wb.getSheetAt(1));
            // 特殊二体合成
            Common.readSpecials(context, wb.getSheetAt(2));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Common.sortMaterialPersonas(context);
    }
}
