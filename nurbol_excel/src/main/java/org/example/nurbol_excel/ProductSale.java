package org.example.nurbol_excel;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ProductSale {
    private int id;
    private String name;
    private double price;
    private int quantity;
    private double finalPrice;
    private String date;
}