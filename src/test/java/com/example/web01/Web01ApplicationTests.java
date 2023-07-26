package com.example.web01;

import com.example.web01.service.TestService;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import javax.annotation.Resource;

@SpringBootTest
class Web01ApplicationTests {

    @Resource
    private TestService testService;

    @Test
    void contextLoads() {
        String s = testService.getHello();
        System.out.println(s);
    }

}
