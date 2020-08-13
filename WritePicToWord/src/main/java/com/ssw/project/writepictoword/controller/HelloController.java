package com.ssw.project.writepictoword.controller;

import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;

/**
 * Created with IntelliJ IDEA.
 *
 * @Auther: ssw
 * @Date: 2020/08/11/18:03
 * @Description:
 */
@Controller
public class HelloController {

    @RequestMapping("/hello")
    public String hello(Model model, HttpServletRequest request) {
        model.addAttribute("msg", "hello,谢谢!");
        request.setAttribute("name","你");
        return "index";
    }

}
