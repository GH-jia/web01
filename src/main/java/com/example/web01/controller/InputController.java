package com.example.web01.controller;

import org.springframework.core.io.ClassPathResource;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.InputStream;
import java.util.UUID;

@RestController
@RequestMapping("/input")
public class InputController {

    @PostMapping("inputWord")
    public String inputWord(@RequestParam MultipartFile file){
        try{
            InputStream inputStream = file.getInputStream();
            String fname = UUID.randomUUID().toString();
            File files = new ClassPathResource("files").getFile();
            File f = new File(files, fname);
        }catch (Exception ex){

        }

        return "success";
    }
}
