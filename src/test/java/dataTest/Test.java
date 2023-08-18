package dataTest;

import utils.Acciones;

import java.util.*;

public class Test {
    public static void main(String[] args) {
        Object[][] array1 = {{ "manzana", 5, "banana", 3, "pera", 10 },
                {"some", 3, "somelse", "any", "anyelse"},
                {"la Banana", "la banana al cuadrado", "la manzana"}};
        Object[][] array2 = {{ 5, "banana", "pera", "manzana", 10, 3 },
                {"la manzana", "la Banana", "la banana al cuadrado"},
                {"somelse", "any", "some", 3, "anyelse"}
                };

        System.out.println(Arrays.deepToString(array1));

        int rows = array1.length;
        System.out.println(rows);

        shuffleArray(array1);
        shuffleArray(array2);
        System.out.println(Arrays.toString(array1) + "-->Array1");
        System.out.println(Arrays.toString(array2) + "-->Array2");

        // Convertir los arrays en listas para comparación
        List<Object> list1 = new ArrayList<>(Arrays.asList(array1));
        List<Object> list2 = new ArrayList<>(Arrays.asList(array2));
        System.out.println(list1 + "-->List 1");
        System.out.println(list2 + "-->List 2");

        // Paso 3: Ordenar las listas
        Collections.sort(list1, new MixedComparators());
        Collections.sort(list2, new MixedComparators());
        System.out.println(list1 + "-->List 1 after");
        System.out.println(list2 + "-->list 2 afyer");


        System.out.println("----------------------------------------------------------------------------------------------");
        for (int i = 0; i < list1.size(); i++) {
            if (array1[i].length < array2.length){
                array1[i] = array1[i+1];
                System.out.println("Array: " + Arrays.deepToString(array1));
            }

        }
        if (list1.size() < list2.size()){
            list1.add("0");
        }else if (list2.size() < list1.size()){
            list2.add("0");
        }
        System.out.println("-------------------------------------------------------------------------------------------");
            for (int k = 0; k < list1.size(); k++) {
                for (int j = 0; j < list1.size() - 1; j++) {
                    if (compareRows(list1, list2)){
                        System.out.println("La fila " + k + " y la fila " + j + " son iguales en ambas matrices.");
                        System.out.println(list1 + "\n" + list2);


                    }
                }
            }


        System.out.println("-------------------------------------------------------------------------------------------------------");

        /*// Paso 1: Desordenar los arrays
        shuffleArray(array1);
        shuffleArray(array2);
        System.out.println(Arrays.toString(array1) + "-->Array1");
        System.out.println(Arrays.toString(array2) + "-->Array2");

        // Convertir los arrays en listas para comparación
        List<Object> list1 = new ArrayList<>(Arrays.asList(array1));
        List<Object> list2 = new ArrayList<>(Arrays.asList(array2));
        System.out.println(list1 + "-->List 1");
        System.out.println(list2 + "-->List 2");

        // Paso 3: Ordenar las listas
        Collections.sort(list1, new MixedComparators());
        Collections.sort(list2, new MixedComparators());
        System.out.println(list1 + "-->List 1 after");
        System.out.println(list2 + "-->list 2 afyer");*/

        // Paso 4: Comparar las listas ordenadas
        boolean sameCombination = list1.equals(list2);

        System.out.println("Los arrays tienen la misma combinación de elementos: " + sameCombination);

    }

    // Función para desordenar el array utilizando el algoritmo de Fisher-Yates
    private static void shuffleArray(Object[] array) {
        for (int i = array.length - 1; i > 0; i--) {
            int j = (int) (Math.random() * (i + 1));
            Object temp = array[i];
            array[i] = array[j];
            array[j] = temp;
        }
    }

    public static boolean compareRows(List<Object> row1, List<Object> row2) {
        if (row1.size() != row2.size()) {
            return false;
        }
        //Arrays.sort(row1);
        //System.out.println("Así se ve la vuelta ordenada: " + Arrays.toString(row1));
        //Arrays.sort(row2);
        //System.out.println("Así se ve la vuelta ordenada: " + Arrays.toString(row2));

        for (int i = 0; i < row1.size(); i++) {
            if (row1.get(i) != row2.get(i)) {
                return false;
            }
        }

        return true;
    }
}

class MixedComparators implements java.util.Comparator<Object> {
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

