package persona;

import java.util.*;

public class Context {
    public Context(String materialFile, String resultFile) {
        this.materialFile = materialFile;
        this.resultFile = resultFile;
    }

    /**
     * 资料源
     */
    public final String materialFile;

    /**
     * 输出结果
     */
    public final String resultFile;

    /**
     * 全部面具（包括宝魔）
     */
    public final List<Persona> personaList = new ArrayList<>();

    /**
     * 全部面具按塔罗牌名称分组，key：塔罗牌名称
     */
    public final TreeMap<String, List<Persona>> personaByArcana = new TreeMap<>();

    /**
     * 全部面具以名称为key
     */
    public final Map<String, Persona> personaByName = new LinkedHashMap<>();

    /**
     * key1 = 战斗面具的塔罗牌，key2 = 宝魔面具，value = 升降几个排名
     */
    public final TreeMap<String, TreeMap<String, Integer>> preciousIncr = new TreeMap<>();

    /**
     * 两个塔罗牌组合会合成哪一种塔罗牌
     */
    public final Map<MaterialPair, String> materialArcanaMap = new LinkedHashMap<>();

    /**
     * 特殊面具合成
     */
    public final Map<MaterialPair, String> materialSpecialPersonaMap = new LinkedHashMap<>();

    /**
     * 输出结果
     */
    public final TreeMap<Persona, TreeMap<Persona, Persona>> result = new TreeMap<>();

}
