package com.example.poi.service;

import com.example.poi.model.Case;
import java.util.List;

public interface ExportService {

  void exportFile(List<Case> cases);
}
