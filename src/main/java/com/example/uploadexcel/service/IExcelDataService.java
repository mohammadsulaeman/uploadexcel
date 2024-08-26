package com.example.uploadexcel.service;

import com.example.uploadexcel.domain.Users;

import java.io.IOException;
import java.util.List;

public interface IExcelDataService {
    List<Users> getExcelDataAsList(String fileName) throws IOException;

    int saveExcelData(List<Users> invoices);
}
