package com.example.easyexcel_showcase.controller;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class Controller {

    @GetMapping("aaa")
    public void aaa(){
        System.out.println();
    }
}
