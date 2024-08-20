package com.mana.exporter;

import lombok.Builder;
import lombok.Data;
import lombok.ToString;

@Data
@Builder
@ToString
public class PersonTest {
    private String name;
    private int age;
}
