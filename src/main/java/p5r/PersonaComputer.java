package p5r;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * Persona 5 Royal 全面具两两配对合成结果计算器。转载代码请保留此处全部注释！
 *
 * @author 陈增辉
 * @date 2024年1月27日
 * @link <a href="https://www.bilibili.com/read/cv19379754/">计算公式</a>
 * @link <a href="https://wiki.biligame.com/persona/P5R/%E4%BA%BA%E6%A0%BC%E9%9D%A2%E5%85%B7%E5%9B%BE%E9%89%B4">人格面具图鉴</a>
 * @link <a href="https://wiki.biligame.com/persona/P5R/%E5%90%88%E6%88%90%E8%8C%83%E5%BC%8F">对照表</a>
 */
public class PersonaComputer {

    /**
     * 资料源
     */
    private static final String MATERIAL_FILE = "persona.xlsx";

    /**
     * 输出结果
     */
    private static final String RESULT_FILE = "result.xlsx";

    /**
     * 全部面具（包括宝魔）
     */
    private static final List<Persona> PERSONA_LIST = new ArrayList<>();

    /**
     * 全部面具按塔罗牌名称分组，key：塔罗牌名称
     */
    private static final TreeMap<String, List<Persona>> PERSONA_BY_ARCANA = new TreeMap<>();

    /**
     * 全部面具以名称为key
     */
    private static final Map<String, Persona> PERSONA_BY_NAME = new LinkedHashMap<>();

    /**
     * key1 = 战斗面具的塔罗牌，key2 = 宝魔面具，value = 升降几个排名
     */
    private static final TreeMap<String, TreeMap<String, Integer>> PRECIOUS_INCR = new TreeMap<>();

    /**
     * 两个塔罗牌组合会合成哪一种塔罗牌
     */
    private static final Map<MaterialPair, String> MATERIAL_ARCANA_MAP = new LinkedHashMap<>();

    /**
     * 特殊面具合成
     */
    private static final Map<MaterialPair, String> MATERIAL_SPECIAL_PERSONA_MAP = new LinkedHashMap<>();

    /**
     * 集体合成的面具
     */
    private static final Set<String> PERSONA_BY_GROUP = Stream.of("义经", "米迦勒", "梅塔特隆", "隐形鬼", "路西法", "撒旦耶尔",
            "佛劳洛斯", "塔姆林", "猫将军", "地狱天使", "赛特", "吹号者", "巴古斯", "邪恶霜精", "婆苏吉", "黄龙", "阿修罗王", "斯拉欧加", "蚩尤"
    ).collect(Collectors.toSet());

    /**
     * 输出结果
     */
    private static final TreeMap<Persona, TreeMap<Persona, Persona>> RESULT = new TreeMap<>();

    public static void main(String[] args) {
        MATERIAL_SPECIAL_PERSONA_MAP.put(new MaterialPair("巴隆", "兰达"), "湿婆");
        MATERIAL_SPECIAL_PERSONA_MAP.put(new MaterialPair("贝利亚", "奈比洛斯"), "爱丽丝");
        MATERIAL_SPECIAL_PERSONA_MAP.put(new MaterialPair("帕尔瓦蒂", "湿婆"), "阿尔达");
        File file = new File(RESULT_FILE);
        if (file.exists() && !file.delete()) {
            System.err.println("请先删除当前目录下的" + RESULT_FILE + "文件！");
            return;
        }
        readPersonas();
        compute();
        System.out.println("debug");
        saveExcel();
    }

    private static void saveExcel(SXSSFWorkbook wb) {
        SXSSFSheet sheet = wb.createSheet("合成表");
        int rownum = 0;
        SXSSFRow row1 = sheet.createRow(rownum++);
        SXSSFRow row2 = sheet.createRow(rownum++);
        int columnIndex = 2;
        for (Persona p2 : RESULT.values().stream().findFirst().get().keySet()) {
            row1.createCell(columnIndex).setCellValue(p2.arcana + "Lv" + p2.level);
            row2.createCell(columnIndex++).setCellValue(p2.name);
        }
        for (Map.Entry<Persona, TreeMap<Persona, Persona>> entry : RESULT.entrySet()) {
            Persona p1 = entry.getKey();
            TreeMap<Persona, Persona> p2AndResultMap = entry.getValue();
            columnIndex = 0;
            SXSSFRow row = sheet.createRow(rownum++);
            row.createCell(columnIndex++).setCellValue(p1.arcana + "Lv" + p1.level);
            row.createCell(columnIndex++).setCellValue(p1.name);
            for (Map.Entry<Persona, Persona> p2AndResult : p2AndResultMap.entrySet()) {
                Persona result = p2AndResult.getValue();
                row.createCell(columnIndex++).setCellValue(result != null ? result.format() : "");
            }
        }
    }

    private static void saveExcel() {
        try (FileOutputStream out = new FileOutputStream(RESULT_FILE);
             SXSSFWorkbook wb = new SXSSFWorkbook()) {
            saveExcel(wb);
            wb.write(out);
            wb.dispose();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static void compute() {
        for (Persona p1 : PERSONA_LIST) {
            String arcana1 = p1.arcana;
            String name1 = p1.name;
            for (Persona p2 : PERSONA_LIST) {
                String arcana2 = p2.arcana;
                String name2 = p2.name;
                Persona result;
                if (name1.equals(name2)) {
                    result = null;
                } else if (p1.isPrecious && p2.isPrecious) {
                    // 两个宝魔，不可合成
                    result = null;
                } else if (p1.isPrecious || p2.isPrecious) {
                    // 有且只有一个宝魔
                    String preciousName; // 宝魔名称
                    String arcana; // 战斗面具的塔罗牌
                    String materialPersona;
                    if (p1.isPrecious) {
                        preciousName = name1;
                        arcana = arcana2;
                        materialPersona = name2;
                    } else {
                        preciousName = name2;
                        arcana = arcana1;
                        materialPersona = name1;
                    }
                    TreeMap<String, Integer> preciousMap = PRECIOUS_INCR.get(arcana);
                    Integer incr;
                    if (preciousMap == null || (incr = preciousMap.get(preciousName)) == null) {
                        result = null;
                    } else {
                        List<Persona> results = PERSONA_BY_ARCANA.get(arcana); // 同一塔罗牌的所有面具
                        int materialIndex = -1;
                        for (int i = 0; i < results.size(); i++) {
                            Persona persona = results.get(i);
                            if (persona.name.equals(materialPersona)) {
                                materialIndex = i;
                                break;
                            }
                        }
                        int targetIndex = materialIndex + incr;
                        result = null;
                        while (targetIndex >= 0 && targetIndex < results.size()) {
                            Persona target = results.get(targetIndex);
                            if (isNotSpecial(target, name1, name2)) {
                                result = target;
                                break;
                            } else {
                                targetIndex = targetIndex + (incr > 0 ? 1 : -1);
                            }
                        }
                    }
                } else {
                    // 没有宝魔
                    // 检查特殊合成
                    Persona checkSpecial = checkSpecial(name1, name2);
                    if (checkSpecial != null) {
                        result = checkSpecial;
                    } else {
                        // 标准合成
                        String resultArcana = MATERIAL_ARCANA_MAP.get(new MaterialPair(arcana1, arcana2));
                        if (resultArcana == null || resultArcana.trim().isEmpty()) {
                            result = null;
                        } else {
                            int level = (p1.level + p2.level) / 2 + 1;
                            List<Persona> results = PERSONA_BY_ARCANA.get(resultArcana);
                            if (arcana1.equals(arcana2)) {
                                // 同种塔罗牌，往低段位找
                                result = findLower(results, level, name1, name2);
                            } else {
                                // 不同种塔罗牌，往高段位找
                                result = results.stream()
                                        .filter(persona -> persona.level >= level && isNotSpecial(persona, name1, name2))
                                        .findFirst()
                                        .orElse(null);
                                if (result == null) {
                                    // 往高段位找不到，就改为往低段位找
                                    result = findLower(results, level, name1, name2);
                                }
                            }
                        }
                    }
                }
                RESULT.computeIfAbsent(p1, p -> new TreeMap<>()).put(p2, result);
            }
        }
    }

    private static Persona findLower(List<Persona> results, int level, String name1, String name2) {
        for (int i = results.size() - 1; i >= 0; i--) {
            Persona persona = results.get(i);
            if (persona.level <= level && isNotSpecial(persona, name1, name2)) {
                return persona;
            }
        }
        return null;
    }

    private static boolean isNotSpecial(Persona persona, String name1, String name2) {
        String name = persona.name;
        return !MATERIAL_SPECIAL_PERSONA_MAP.containsValue(name) && !PERSONA_BY_GROUP.contains(name)
                && !name.equals(name1) && !name.equals(name2)
                && !persona.isPrecious;
    }

    private static Persona checkSpecial(String name1, String name2) {
        String result = MATERIAL_SPECIAL_PERSONA_MAP.get(new MaterialPair(name1, name2));
        if (result == null) {
            result = MATERIAL_SPECIAL_PERSONA_MAP.get(new MaterialPair(name2, name1));
        }
        return result == null ? null : PERSONA_BY_NAME.get(result);
    }

    private static void readPersonas() {
        try (XSSFWorkbook wb = new XSSFWorkbook(MATERIAL_FILE)) {
            // 素材面具
            XSSFSheet sheet = wb.getSheetAt(0);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    // 跳过标题行
                    continue;
                }
                String arcana = row.getCell(0).getStringCellValue();
                int level = Double.valueOf(row.getCell(1).getNumericCellValue()).intValue();
                String name = row.getCell(2).getStringCellValue();
                Cell cell = row.getCell(3);
                boolean isPrecious = cell != null && "宝魔".equals(cell.getStringCellValue());
                Persona persona = new Persona(arcana, level, name, isPrecious);
                PERSONA_LIST.add(persona);
                PERSONA_BY_NAME.put(persona.name, persona);
                PERSONA_BY_ARCANA.computeIfAbsent(persona.arcana, a -> new ArrayList<>()).add(persona);
            }
            // 宝魔升降
            sheet = wb.getSheetAt(1);
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
                        PRECIOUS_INCR.computeIfAbsent(arcana, p -> new TreeMap<>()).put(preciousName, incr);
                    }
                }
            }
            // 塔罗牌二合一
            sheet = wb.getSheetAt(2);
            Map<Integer, String> rowMap = new TreeMap<>(); // key = columnIndex
            for (Row row : sheet) {
                String arcana1 = "";
                for (Cell cell : row) {
                    if (row.getRowNum() == 0) {
                        int columnIndex = cell.getColumnIndex();
                        if (columnIndex != 0) {
                            rowMap.put(columnIndex, cell.getStringCellValue());
                        }
                    } else if (cell.getColumnIndex() == 0) {
                        arcana1 = cell.getStringCellValue();
                    } else {
                        String arcana2 = rowMap.get(cell.getColumnIndex());
                        String result = cell.getStringCellValue();
                        MATERIAL_ARCANA_MAP.put(new MaterialPair(arcana1, arcana2), result);
                    }
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static class Persona implements Comparable<Persona> {
        final String arcana;
        final Integer level;
        final String name;
        /**
         * 是否宝魔
         */
        final boolean isPrecious;

        public Persona(String arcana, Integer level, String name, boolean isPrecious) {
            this.arcana = arcana;
            this.level = level;
            this.name = name;
            this.isPrecious = isPrecious;
        }

        public String format() {
            return arcana + "Lv" + level + " " + name;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Persona persona = (Persona) o;
            return isPrecious == persona.isPrecious && Objects.equals(arcana, persona.arcana) && Objects.equals(level, persona.level) && Objects.equals(name, persona.name);
        }

        @Override
        public int hashCode() {
            return Objects.hash(arcana, level, name, isPrecious);
        }

        @Override
        public String toString() {
            return "Persona{" +
                    "arcana='" + arcana + '\'' +
                    ", level=" + level +
                    ", name='" + name + '\'' +
                    ", isPrecious=" + isPrecious +
                    '}';
        }

        @Override
        public int compareTo(Persona o) {
            int compare = this.level.compareTo(o.level);
            if (compare != 0) {
                return compare;
            }
            return this.name.compareTo(o.name);
        }
    }

    private static class MaterialPair {
        /**
         * 素材1
         */
        final String m1;
        /**
         * 素材2
         */
        final String m2;

        public MaterialPair(String m1, String m2) {
            this.m1 = m1;
            this.m2 = m2;
        }

        @Override
        public String toString() {
            return "MaterialPair{" +
                    "m1='" + m1 + '\'' +
                    ", m2='" + m2 + '\'' +
                    '}';
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            MaterialPair that = (MaterialPair) o;
            return Objects.equals(m1, that.m1) && Objects.equals(m2, that.m2);
        }

        @Override
        public int hashCode() {
            return Objects.hash(m1, m2);
        }
    }
}
