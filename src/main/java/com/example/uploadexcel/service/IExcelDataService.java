package com.example.uploadexcel.service;

import com.example.uploadexcel.domain.Users;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

public interface IExcelDataService {
    List<Users> getExcelDataAsList() throws IOException;

    int saveExcelData(List<Users> invoices);
}
