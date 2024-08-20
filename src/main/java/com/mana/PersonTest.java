package com.mana;

import lombok.Builder;
import lombok.Data;
import lombok.ToString;

import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Data
@Builder
@ToString
public class PersonTest {
    private String name;
    private Integer age;
    private Date birthDate;
    private Boolean alive;
    private Double salary;
    private Gender gender;
    private String[] colors;
    private Integer[] sonAges;
    private List<String> hobbies;
    private Set<Double> marks;
    private Map<String, Integer> course;
}
