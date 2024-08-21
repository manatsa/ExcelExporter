package com.mana;

import lombok.Builder;
import lombok.Data;
import lombok.ToString;

/**
 * @author :: codemaster
 * created on :: 21/8/2024
 * Package Name :: com.mana
 */

@Data
@Builder
@ToString
public class Address {
    private String street;
    private String city;
    private String country;
}
