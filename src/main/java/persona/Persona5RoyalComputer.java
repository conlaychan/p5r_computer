package persona;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;

/**
 * Persona 5 Royal 全面具两两配对合成结果计算器。转载代码请保留此处全部注释！
 *
 * @author 陈增辉 2024年1月27日
 * @link <a href="https://www.bilibili.com/read/cv19379754/">计算公式</a>
 * @link <a href="https://wiki.biligame.com/persona/P5R/%E4%BA%BA%E6%A0%BC%E9%9D%A2%E5%85%B7%E5%9B%BE%E9%89%B4">人格面具图鉴</a>
 * @link <a href="https://wiki.biligame.com/persona/P5R/%E5%90%88%E6%88%90%E8%8C%83%E5%BC%8F">对照表</a>
 */
public class Persona5RoyalComputer {
    public static void computer() {
        System.out.println("Persona 5 Royal：");

        Context context = new Context("Persona5Royal_material.xlsx", "Persona5Royal_computed.xlsx");
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

            // 宝魔升降
            XSSFSheet sheet = wb.getSheetAt(1);
            Map<Integer, String> arcanaMap = new TreeMap<>();
            for (Cell cell : sheet.getRow(0)) {
                int columnIndex = cell.getColumnIndex();
                if (columnIndex != 0) {
                    arcanaMap.put(columnIndex, cell.getStringCellValue());
                }
            }
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                String preciousName = ""; // 宝魔面具
                for (Cell cell : row) {
                    int columnIndex = cell.getColumnIndex();
                    if (columnIndex == 0) {
                        preciousName = cell.getStringCellValue();
                    } else {
                        int incr = Double.valueOf(cell.getNumericCellValue()).intValue();
                        String arcana = arcanaMap.get(columnIndex); // 战斗面具的塔罗牌
                        context.preciousIncr.computeIfAbsent(arcana, p -> new TreeMap<>()).put(preciousName, incr);
                    }
                }
            }
            // 塔罗牌二合一
            Common.readArcanaRules(context, wb.getSheetAt(2));
            // 特殊二体合成
            Common.readSpecials(context, wb.getSheetAt(3));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        Common.sortMaterialPersonas(context);
    }
}
