package persona;

import java.util.Objects;

public class Persona implements Comparable<Persona> {
    public final String arcana;
    public final Integer level;
    public final String name;
    /**
     * 是否宝魔
     */
    public final boolean isPrecious;

    /**
     * 需要集体断头台合成
     */
    public final boolean isGroup;

    public Persona(String arcana, Integer level, String name, boolean isPrecious, boolean isGroup) {
        this.arcana = arcana;
        this.level = level;
        this.name = name;
        this.isPrecious = isPrecious;
        this.isGroup = isGroup;
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