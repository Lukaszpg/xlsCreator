import commons.XlsFileCreator;
import model.Person;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        XlsFileCreator<Person> xlsFileCreator = new XlsFileCreator<>(Person.class);

        List<Person> persons = Arrays.asList(
                new Person("Adam", "Kowalski", 25),
                new Person("Maria", "Kowalska", 22));

        try {
            xlsFileCreator.createFile(persons, "src/main/resources/", "persons");

        } catch (NoSuchMethodException | InvocationTargetException | IllegalAccessException | IOException e) {
            e.printStackTrace();
        }

    }
}

