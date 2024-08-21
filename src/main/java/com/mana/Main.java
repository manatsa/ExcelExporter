package com.mana;

import com.mana.exporter.ExcelExporter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.*;

public class Main {
    public static void main(String[] args) throws Exception {

        var exporter=new ExcelExporter();
        PersonTest mana=PersonTest
                .builder()
                .name("Manatsa")
                .alive(Boolean.TRUE)
                .birthDate(new Date(1986,10,24))
                .salary(1000.01)
                .gender(Gender.MALE)
                .colors(new String[]{"Black","Navy Blue","White"})
                .sonAges(new Integer[]{3,5,8})
                .marks(Set.of(86.75,90.45,37.0))
                .hobbies(List.of("Singing", "Walking", "Programming"))
                .course(Map.of("Maths",2010,"Comp Science",2012))
                .age(38)
                .build();
        PersonTest grace= PersonTest
                .builder()
                .name("Grace")
                .birthDate(new Date(1986,11,18))
                .alive(Boolean.TRUE)
                .salary(620.14)
                .colors(new String[]{"Blue","White","Brown"})
                .sonAges(new Integer[]{3,5,8})
                .gender(Gender.FEMALE)
                .marks(Set.of(56.05,70.0,57.0))
                .hobbies(List.of("Working", "Telling Stories", "Relaxing"))
                .course(Map.of("Divorce",2023,"Abuse",2024))
                .age(30)
                .build();
        List<PersonTest> personTestList=List.of(mana,grace);
        FileOutputStream fileOutputStream=new FileOutputStream(new File("TestExcel.xlsx"));

        exporter.exportToExcel(personTestList, fileOutputStream, new XSSFWorkbook());
    }
}