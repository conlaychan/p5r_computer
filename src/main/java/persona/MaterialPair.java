package persona;

import java.util.Objects;

public class MaterialPair {
    /**
     * 素材1
     */
    public final String m1;
    /**
     * 素材2
     */
    public final String m2;

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