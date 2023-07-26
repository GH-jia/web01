package com.example.web01;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("com.example.web01.mapper")
public class Web01Application {

    public static void main(String[] args) {
        SpringApplication.run(Web01Application.class, args);
    }

}
