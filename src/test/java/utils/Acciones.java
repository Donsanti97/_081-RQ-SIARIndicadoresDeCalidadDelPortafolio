package utils;

import java.util.Comparator;
import java.util.Objects;

public class Acciones {
    public class MixedComparators implements java.util.Comparator<Object> {

        @Override
        public int compare(Object o1, Object o2) {
            if (o1 instanceof String && o2 instanceof String) {
                return ((String) o1).compareTo((String) o2);
            } else if (o1 instanceof Integer && o2 instanceof Integer) {
                return Integer.compare((Integer) o1, (Integer) o2);
            } else if (o1 instanceof String) {
                return -1; // Cadenas antes que enteros
            } else if (o2 instanceof String) {
                return 1;  // Cadenas antes que enteros
            } else {
                return 0;  // Tipos desconocidos (no debería ocurrir)
            }
        }


    }

}


