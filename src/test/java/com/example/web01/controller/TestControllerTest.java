package com.example.web01.controller;

import com.baomidou.mybatisplus.core.conditions.Wrapper;
import com.baomidou.mybatisplus.core.conditions.query.LambdaQueryWrapper;
import com.baomidou.mybatisplus.core.conditions.query.QueryWrapper;
import com.baomidou.mybatisplus.core.conditions.segments.MergeSegments;
import com.example.web01.mapper.UserMapper;
import com.example.web01.pojo.User;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.util.Assert;

import javax.annotation.Resource;

import java.util.List;

@SpringBootTest
class TestControllerTest {

    @Resource
    UserMapper userMapper;

    @Test
    void testMapper(){
        List<User> userList = userMapper.selectList(null);
        Assert.isTrue(5 == userList.size(),"");
        userList.forEach(System.out::println);
        QueryWrapper<User> wrapper = new QueryWrapper<>();
        wrapper.eq("name","Jone");
        User user = userMapper.selectOne(wrapper);
        System.out.println(user.toString());
    }


}
