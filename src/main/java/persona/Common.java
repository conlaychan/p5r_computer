package persona;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Common {

    public static void saveExcel(SXSSFWorkbook wb, Context context) {
        SXSSFSheet sheet = wb.createSheet("合成表");
        SXSSFSheet sheet2 = wb.createSheet("反向查询");
        AuthorDescription.append(wb);
        int rownum = 0;
        SXSSFRow row1 = sheet.createRow(rownum++);
        SXSSFRow row2 = sheet.createRow(rownum++);
        int columnIndex = 2;
        for (Persona p2 : context.result.values().stream().findFirst().orElse(new TreeMap<>()).keySet()) {
            row1.createCell(columnIndex).setCellValue(p2.arcana + "Lv" + p2.level);
            row2.createCell(columnIndex++).setCellValue(p2.name);
        }
        for (Map.Entry<Persona, TreeMap<Persona, Persona>> entry : context.result.entrySet()) {
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

        // 反向查询
        saveInverse(context, sheet2);
    }

    private static void saveInverse(Context context, SXSSFSheet sheet) {
        // k = 合成结果，v = 合成材料
        TreeMap<Persona, Set<MaterialPair>> inverseHint = new TreeMap<>();
        context.result.forEach((p1, p2AndResultMap) -> {
            p2AndResultMap.forEach((p2, result) -> {
                if (result != null) {
                    Set<MaterialPair> materialPairs = inverseHint.computeIfAbsent(result, k -> new HashSet<>());
                    MaterialPair pair1 = new MaterialPair(p1.name, p2.name);
                    MaterialPair pair2 = new MaterialPair(p2.name, p1.name);
                    if (!materialPairs.contains(pair1) && !materialPairs.contains(pair2)) {
                        materialPairs.add(pair1);
                    }
                }
            });
        });
        // k = 合成结果，v = 合成材料
        TreeMap<Persona, List<MaterialPersonaPair>> inverse = new TreeMap<>();
        inverseHint.forEach((result, materialPairs) -> {
            List<MaterialPersonaPair> materils = new ArrayList<>();
            for (MaterialPair materialPair : materialPairs) {
                Persona p1 = context.personaByName.get(materialPair.m1);
                Persona p2 = context.personaByName.get(materialPair.m2);
                materils.add(new MaterialPersonaPair(p1, p2));
            }
            materils.sort(Comparator.comparingInt(a -> Math.max(a.p1.level, a.p2.level)));
            inverse.put(result, materils);
        });
        // 开始写入
        int resultIndex = 0;
        int maxSize = 0;
        SXSSFRow row = sheet.createRow(0);
        for (Map.Entry<Persona, List<MaterialPersonaPair>> entry : inverse.entrySet()) {
            Persona result = entry.getKey();
            int size = entry.getValue().size();
            if (size > maxSize) {
                maxSize = size;
            }
            int colIndex = resultIndex * 2;
            sheet.addMergedRegion(new CellRangeAddress(0, 0, colIndex, colIndex + 1));
            row.createCell(colIndex).setCellValue(result.format());
            resultIndex++;
        }
        for (int pairIndex = 0; pairIndex < maxSize; pairIndex++) {
            resultIndex = 0;
            for (Map.Entry<Persona, List<MaterialPersonaPair>> entry : inverse.entrySet()) {
                List<MaterialPersonaPair> pairs = entry.getValue();
                if (pairIndex < pairs.size()) {
                    MaterialPersonaPair pair = pairs.get(pairIndex);
                    row = sheet.getRow(pairIndex + 1);
                    if (row == null) {
                        row = sheet.createRow(pairIndex + 1);
                    }
                    SXSSFCell cell = row.createCell(resultIndex * 2);
                    cell.setCellValue(pair.p1.format());
                    cell = row.createCell(resultIndex * 2 + 1);
                    cell.setCellValue(pair.p2.format());
                }
                resultIndex++;
            }
        }
    }

    private static class MaterialPersonaPair {
        final Persona p1;
        final Persona p2;

        private MaterialPersonaPair(Persona p1, Persona p2) {
            this.p1 = p1;
            this.p2 = p2;
        }
    }

    public static void saveExcel(Context context) {
        try (FileOutputStream out = new FileOutputStream(context.resultFile);
             SXSSFWorkbook wb = new SXSSFWorkbook()) {
            saveExcel(wb, context);
            wb.write(out);
            wb.dispose();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void sortMaterialPersonas(Context context) {
        context.personaList.sort(Persona::compareTo);
        for (Persona persona : context.personaList) {
            context.personaByName.put(persona.name, persona);
            context.personaByArcana.computeIfAbsent(persona.arcana, a -> new ArrayList<>()).add(persona);
        }
    }

    public static void readSpecials(Context context, XSSFSheet sheet) {
        for (Row row : sheet) {
            if (row.getRowNum() != 0) {
                String m1 = row.getCell(0).getStringCellValue();
                String m2 = row.getCell(1).getStringCellValue();
                String result = row.getCell(2).getStringCellValue();
                context.materialSpecialPersonaMap.put(new MaterialPair(m1, m2), result);
            }
        }
    }

    public static void readArcanaRules(Context context, XSSFSheet sheet) {
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
                    context.materialArcanaMap.put(new MaterialPair(arcana1, arcana2), result);
                }
            }
        }
    }

    public static void readMaterialPersonas(Context context, XSSFSheet sheet) {
        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                // 跳过标题行
                continue;
            }
            String arcana = row.getCell(0).getStringCellValue();
            int level = Double.valueOf(row.getCell(1).getNumericCellValue()).intValue();
            String name = row.getCell(2).getStringCellValue();
            Cell cell = row.getCell(3);
            boolean isPrecious;
            boolean isGroup;
            if (cell != null) {
                String cellValue = cell.getStringCellValue();
                isPrecious = "宝魔".equals(cellValue);
                isGroup = "集体合成".equals(cellValue);
            } else {
                isPrecious = false;
                isGroup = false;
            }
            Persona persona = new Persona(arcana, level, name, isPrecious, isGroup);
            context.personaList.add(persona);
        }
    }

    public static Persona findLower(List<Persona> results, int level, String name1, String name2, Context context) {
        for (int i = results.size() - 1; i >= 0; i--) {
            Persona persona = results.get(i);
            if (persona.level <= level && isNotSpecial(persona, name1, name2, context)) {
                return persona;
            }
        }
        return null;
    }

    public static boolean isNotSpecial(Persona persona, String name1, String name2, Context context) {
        String name = persona.name;
        return !context.materialSpecialPersonaMap.containsValue(name) && !persona.isGroup
                && !name.equals(name1) && !name.equals(name2)
                && !persona.isPrecious;
    }

    public static Persona checkSpecial(String name1, String name2, Context context) {
        String result = context.materialSpecialPersonaMap.get(new MaterialPair(name1, name2));
        if (result == null) {
            result = context.materialSpecialPersonaMap.get(new MaterialPair(name2, name1));
        }
        return result == null ? null : context.personaByName.get(result);
    }

    public static void compute(Context context) {
        for (Persona p1 : context.personaList) {
            String arcana1 = p1.arcana;
            String name1 = p1.name;
            for (Persona p2 : context.personaList) {
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
                    TreeMap<String, Integer> preciousMap = context.preciousIncr.get(arcana);
                    Integer incr;
                    if (preciousMap == null || (incr = preciousMap.get(preciousName)) == null) {
                        result = null;
                    } else {
                        List<Persona> results = context.personaByArcana.get(arcana); // 同一塔罗牌的所有面具
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
                            if (Common.isNotSpecial(target, name1, name2, context)) {
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
                    Persona checkSpecial = Common.checkSpecial(name1, name2, context);
                    if (checkSpecial != null) {
                        result = checkSpecial;
                    } else {
                        // 标准合成
                        String resultArcana = context.materialArcanaMap.get(new MaterialPair(arcana1, arcana2));
                        if (resultArcana == null || resultArcana.trim().isEmpty()) {
                            result = null;
                        } else {
                            int level = (p1.level + p2.level) / 2 + 1;
                            List<Persona> results = context.personaByArcana.get(resultArcana);
                            if (arcana1.equals(arcana2)) {
                                // 同种塔罗牌，往低段位找
                                result = Common.findLower(results, level, name1, name2, context);
                            } else {
                                // 不同种塔罗牌，往高段位找
                                result = results.stream()
                                        .filter(persona -> persona.level >= level && Common.isNotSpecial(persona, name1, name2, context))
                                        .findFirst()
                                        .orElse(null);
                                if (result == null) {
                                    // 往高段位找不到，就改为往低段位找
                                    result = Common.findLower(results, level, name1, name2, context);
                                }
                            }
                        }
                    }
                }
                context.result.computeIfAbsent(p1, p -> new TreeMap<>()).put(p2, result);
            }
        }
    }
}
