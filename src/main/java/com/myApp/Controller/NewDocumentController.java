package com.myApp.Controller;

import com.myApp.Service.DocumentService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

@RequestMapping("/")
@RestController
public class NewDocumentController {
@Autowired
private DocumentService documentService;
    @RequestMapping(method = RequestMethod.GET)
    public @ResponseBody String home() {
        return "Hello World!";
    }

    @RequestMapping(method = RequestMethod.POST, value="username/{name}/{surname}")
    public @ResponseBody String home2(@PathVariable("name") String name, @PathVariable("surname") String surname) {
        documentService.create(name, surname);
        return "Hello," +name+" "+surname;
}}
