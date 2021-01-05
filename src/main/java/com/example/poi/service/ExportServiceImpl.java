package com.example.poi.service;

import com.example.poi.model.Case;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.time.Instant;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import javax.imageio.ImageIO;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

@Service
public class ExportServiceImpl implements ExportService {

  private static final Logger log = LoggerFactory.getLogger(ExportServiceImpl.class);

  @Override
  public void exportFile(List<Case> cases) {
    SXSSFWorkbook workbook = new SXSSFWorkbook(1);
    SXSSFSheet sheet = workbook.createSheet();

    writeCasesToSheet(cases, sheet);
    File file = new File(Instant.now().toEpochMilli() + ".xlsx");

    try (FileOutputStream fileOutputStream = new FileOutputStream(file)) {
      workbook.write(fileOutputStream);
    } catch (IOException e) {
      log.error("Write workbook to byte array output stream error", e);
    }
  }

  private void writeCasesToSheet(List<Case> cases, SXSSFSheet sheet) {
    CellStyle contentRowStyle = sheet.getWorkbook().createCellStyle();
    // 设置成垂直居中
    contentRowStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    contentRowStyle.setWrapText(true);

    int size = cases.size();
    for (int i = 0; i < size; i++) {
      log.debug("Exporting......... {} / {}", i, size);
      SXSSFRow row = sheet.createRow(i);
      row.setRowStyle(contentRowStyle);
      row.setHeightInPoints(10 * row.getHeightInPoints());

      Case ca = cases.get(i);

      row.createCell(0).setCellValue(ca.getName());
      row.createCell(1).setCellValue(ca.getNumber());
      addAttachmentsToSheet(ca.getAttachmentUrls(), row, 2);
    }

  }

  public void addAttachmentsToSheet(List<String> attachments, Row row, int beginIndex) {

    for (int i = 0; i < attachments.size(); i++) {
      Cell cell = row.createCell(beginIndex + i);
      addAttachment(attachments.get(i), cell, i);
    }
  }

  public void addAttachment(String url, Cell cell, int i) {
    if (i % 2 == 0) {

      addHyperLinkToSheet(url, cell);
    } else {
      addPictureToSheet(url, cell);
    }
  }

  public void addHyperLinkToSheet(String url, Cell cell) {
    Workbook workbook = cell.getSheet().getWorkbook();

    CreationHelper createHelper = workbook.getCreationHelper();

    CellStyle hyperLinkStyle = workbook.createCellStyle();
    Font hyperLinkFont = workbook.createFont();
    hyperLinkFont.setUnderline(Font.U_SINGLE);
    hyperLinkFont.setColor(IndexedColors.BLUE.getIndex());
    hyperLinkStyle.setFont(hyperLinkFont);

    Hyperlink hyperlink = createHelper.createHyperlink(HyperlinkType.URL);
    try {
      hyperlink.setAddress(url);
    } catch (Exception e) {
      log.error("Error attachment url is {}", url);
      e.printStackTrace();
    }
    hyperlink.setLabel("Hyper Link");
    cell.setCellValue("Hyper Link");
    cell.setHyperlink(hyperlink);
    cell.setCellStyle(hyperLinkStyle);
  }


  public void addPictureToSheet(String url, Cell cell) {
    try {
      // read image from url
      // could return null if the image not supported by image readers
      BufferedImage originImage = ImageIO.read(new URL(url));
      if (Objects.nonNull(originImage)) {
        double imageWidth = originImage.getWidth();
        double imageHeight = originImage.getHeight();

        addPicture(cell, url, originImage)
            .ifPresent(picIndex -> pictureToSheet(cell, picIndex, imageHeight, imageWidth));
      }
    } catch (IOException e) {
      log.error("Read attachment picture error, attachment url is {}", url, e);
    }
  }

  public Optional<Integer> addPicture(
      Cell cell, String url, BufferedImage originImage) {
    Integer index = null;

    String extension = "png";
    try (ByteArrayOutputStream os = new ByteArrayOutputStream()) {
      ImageIO.write(originImage, extension, os);

      byte[] imageBytes = os.toByteArray();

      // add image to workbook
      index = cell.getSheet().getWorkbook().addPicture(imageBytes, Workbook.PICTURE_TYPE_JPEG);

    } catch (IOException e) {
      log.error("This picture compress error, attachment url {}", url, e);
    }

    return Optional.ofNullable(index);
  }

  private void pictureToSheet(Cell cell, int pictureIdx, double imageHeight, double imageWidth) {
    Drawing patriarch = cell.getSheet().createDrawingPatriarch();
    CreationHelper helper = cell.getSheet().getWorkbook().getCreationHelper();
    ClientAnchor anchor = helper.createClientAnchor();

    anchor.setCol1(cell.getColumnIndex());
    anchor.setRow1(cell.getRow().getRowNum());

    double cellWidth = cell.getSheet().getColumnWidthInPixels(cell.getColumnIndex());
    double cellHeight = cell.getRow().getHeightInPoints() / 72 * 96;
    log.trace("Picture cell height is {}", cellHeight);
    Picture pict = patriarch.createPicture(anchor, pictureIdx);

    if (cellWidth <= imageWidth || cellHeight <= imageHeight) {
      resizePicture(imageHeight, imageWidth, cellWidth, cellHeight, pict);
    } else {
      pict.resize();
    }
  }

  private void resizePicture(
      double imageHeight, double imageWidth, double cellWidth, double cellHeight, Picture pict) {
    if (imageHeight >= imageWidth) {
      double scaleX = (cellHeight * (imageWidth / imageHeight)) / cellWidth;
      log.trace("X scale {}", scaleX);
      pict.resize(scaleX, 1);
    } else {
      double scaleY = (cellWidth * (imageHeight / imageWidth)) / cellHeight;
      scaleY = scaleY * 10 / 4;
      log.trace("Y scale {}", scaleY);
      pict.resize(1, scaleY);
    }
  }
}

