package com.example.poi.model;

import java.util.List;

public class Case {

  private String number;
  private String name;
  private List<String> attachmentUrls;

  public String getNumber() {
    return number;
  }

  public void setNumber(String number) {
    this.number = number;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public List<String> getAttachmentUrls() {
    return attachmentUrls;
  }

  public void setAttachmentUrls(List<String> attachmentUrls) {
    this.attachmentUrls = attachmentUrls;
  }
}
