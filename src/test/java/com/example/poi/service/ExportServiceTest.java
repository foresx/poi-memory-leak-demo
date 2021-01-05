package com.example.poi.service;

import com.example.poi.model.Case;
import com.github.javafaker.Faker;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class ExportServiceTest {

  private final static Faker faker = new Faker();
  private static List<Case> cases = new ArrayList<>(128);
  @Autowired
  private ExportService exportService;

  @BeforeAll
  public static void mockCases() {
    for (int i = 0; i < 128; i++) {
      cases.add(mockCase());
    }
  }

  private static Case mockCase() {
    Case ca = new Case();
    ca.setName(faker.rickAndMorty().character());
    ca.setNumber(faker.number().digits(10));
    List<String> attachmentUrls = new ArrayList<>(10);
    for (int i = 0; i < 10; i++) {
      attachmentUrls.add(
          "https://www.regalhotel.com/uploads/rhi/regal_quarantine-package-720x475px-en.jpg");
    }
    ca.setAttachmentUrls(attachmentUrls);
    return ca;
  }

  @Test
  void exportFile() {
    exportService.exportFile(cases);
  }
}