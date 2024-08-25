package com.example.uploadexcel.controller;

import com.example.uploadexcel.domain.Users;
import com.example.uploadexcel.service.IExcelDataService;
import com.example.uploadexcel.service.IFileUploaderService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import java.io.IOException;
import java.util.List;

@Slf4j
@Controller
public class UploadExcelController {

    @Autowired
    IFileUploaderService fileUploaderService;

    @Autowired
    IExcelDataService excelDataService;

    @PostMapping("/uploadFile")
    public ModelAndView uploadFile(@RequestParam(name = "file") MultipartFile file, Model model) {
        log.info("uploadFile");
        fileUploaderService.uploadFile(file);

        model.addAttribute("message", "You have successfully uploaded '"+ file.getOriginalFilename()+"' !");
        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        return new ModelAndView("index");
    }

    @GetMapping("/saveData")
    public ModelAndView saveData(Model model) throws IOException {
        List<Users> usersList = excelDataService.getExcelDataAsList();
        int noOfRecords  = excelDataService.saveExcelData(usersList);
        model.addAttribute("message","Save Upload File Excel Successfully !"+"jumlah row : "+noOfRecords);
        return new ModelAndView("index");
    }
}
