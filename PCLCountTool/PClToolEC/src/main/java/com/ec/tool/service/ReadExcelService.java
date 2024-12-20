/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.ec.tool.service;

import com.ec.tool.model.NEILModel;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Iterator;
import java.util.List;
import javax.swing.JOptionPane;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author thangnq
 */
public class ReadExcelService {

    public static String fileName;
    public static String fileCode;

    public static NEILModel counting(List<Path> listPath, List<String> tcs) {
        // add variable coutN, countE, countI, countL
        int countN = 0;
        int countE = 0;
        int countI = 0;
        int countL = 0;
        for (final Path path : listPath) {
            if (path.toFile().getName().startsWith("~$")) {
                continue;
            }
            try (final Workbook workbook = new XSSFWorkbook(new FileInputStream(path.toFile()))) {
                final Iterator<Sheet> sheetIterator = workbook.sheetIterator();
                while (sheetIterator.hasNext()) {
                    final Sheet sheet = sheetIterator.next();
                    final String sheetName = sheet.getSheetName();
                    if (!StringUtils.equals(sheetName, "変更履歴") && !StringUtils.contains(sheetName, "修正前")) {
                        // Lặp qua các dòng trong sheet
                        for (int i = 0; i < sheet.getLastRowNum(); i++) {
                            Row row = sheet.getRow(i);
                            if (row != null) {
                                Cell firstColumn = row.getCell(0);
                                if (firstColumn != null && firstColumn.getStringCellValue().equals("ＰＣＬ区分\n"
                                        + "（Ｎ:正常,Ｅ:異常,Ｌ:境界･限界,Ｉ:ｲﾝﾀｰﾌｪｰｽ）")) {
                                    for (int j = 0; j < row.getLastCellNum(); j++) {
                                        Cell cell = row.getCell(j);
                                        if (cell != null) {
                                            if (cell.toString().equals("N")) {
                                                countN++;
                                                String nameTc = sheet.getRow(3).getCell(j).getStringCellValue() + 
                                                        String.format("%03d", ((int) sheet.getRow(4).getCell(j).getNumericCellValue()));
                                                tcs.add(nameTc);
                                            } else if (cell.toString().equals("I")) {
                                                countI++;
                                                String nameTc = sheet.getRow(3).getCell(j).getStringCellValue() + 
                                                        String.format("%03d", ((int) sheet.getRow(4).getCell(j).getNumericCellValue()));
                                                tcs.add(nameTc);
                                            } else if (cell.toString().equals("E")) {
                                                countE++;
                                                String nameTc = sheet.getRow(3).getCell(j).getStringCellValue() + 
                                                        String.format("%03d", ((int) sheet.getRow(4).getCell(j).getNumericCellValue()));
                                                tcs.add(nameTc);
                                            } else if (cell.toString().equals("L")) {
                                                countL++;
                                                String nameTc = sheet.getRow(3).getCell(j).getStringCellValue() + 
                                                        String.format("%03d", ((int) sheet.getRow(4).getCell(j).getNumericCellValue()));
                                                tcs.add(nameTc);
                                            }
                                        }
                                    }
                                }
                                int cellCount = row.getPhysicalNumberOfCells();
                                for (int j = 0; j < cellCount; j++) {
                                    Cell cell = row.getCell(j);
                                    if (cell != null) {
                                        if (StringUtils.equals(String.valueOf(cell.getCellType()), "STRING")) {
                                            if (StringUtils.equals(cell.getStringCellValue(), "プログラムＩＤ")) {
                                                fileCode = row.getCell(j + 5).getStringCellValue();
                                            }

                                            if (StringUtils.equals(cell.getStringCellValue(), "プログラム名")) {
                                                fileName = row.getCell(j + 5).getStringCellValue();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            } catch (IOException e) {
                JOptionPane.showMessageDialog(null, e.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
            }
        }
        return new NEILModel(countN, countE, countI, countL);
    }

    public static void writing(List<Path> listPath, NEILModel nEILModel, String fileOutputPathText) {
        for (final Path path : listPath) {
            if (StringUtils.startsWith(path.toFile().getName(), "D3.1.3 単体テスト品質データ総括表")) {
                writingThemD3(path, nEILModel, fileOutputPathText);
            } else if (StringUtils.startsWith(path.toFile().getName(), "D5.1 簡易バグ管理表")) {
                writingThemD5_1(path, fileOutputPathText);
            } else if (StringUtils.startsWith(path.toFile().getName(), "D5.2 単体テスト(エビデンス)")) {
                writingThemD5_2(path, fileOutputPathText);
            }
        }
    }

    public static void writingThemD3(Path path, NEILModel nEILModel, String fileOutputPathText) {
        try (final Workbook workbook = new XSSFWorkbook(new FileInputStream(path.toFile()))) {
            final Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                final Sheet sheet = sheetIterator.next();
                // Lặp qua các dòng trong sheet
                for (int i = 0; i < sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        int cellCount = row.getPhysicalNumberOfCells();
                        for (int j = 0; j < cellCount; j++) {
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                if (StringUtils.equals(String.valueOf(cell.getCellType()), "STRING")) {
                                    if (StringUtils.equals(cell.getStringCellValue(), "プログラム名称")) {
                                        row.getCell(j + 3).setCellValue(fileCode + "_" + fileName);
                                    }

                                    if (StringUtils.equals(cell.getStringCellValue(), "総 Ｐ Ｃ Ｌ 数")) {
                                        row.getCell(j + 10).setCellValue(nEILModel.getTotal());
                                    }

                                    if (StringUtils.equals(cell.getStringCellValue(), "正 常 項 目(Ｎ)")) {
                                        row.getCell(j + 7).setCellValue(nEILModel.getCountN());
                                    }

                                    if (StringUtils.equals(cell.getStringCellValue(), "異 常 項 目(Ｅ)")) {
                                        row.getCell(j + 7).setCellValue(nEILModel.getCountE());
                                    }

                                    if (StringUtils.equals(cell.getStringCellValue(), "境界・限界項目(Ｌ)")) {
                                        row.getCell(j + 7).setCellValue(nEILModel.getCountL());
                                    }

                                    if (StringUtils.equals(cell.getStringCellValue(), "インターフェース項目(Ｉ)")) {
                                        row.getCell(j + 7).setCellValue(nEILModel.getCountI());
                                    }
                                }
                            }
                        }
                    }
                }
            }
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            exportFileExcel(workbook, path, fileOutputPathText);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, e.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void writingThemD5_1(Path path, String fileOutputPathText) {
        try (final Workbook workbook = new XSSFWorkbook(new FileInputStream(path.toFile()))) {
            final Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                final Sheet sheet = sheetIterator.next();
                // Lặp qua các dòng trong sheet
                for (int i = 0; i < sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        int cellCount = row.getPhysicalNumberOfCells();
                        for (int j = 0; j < cellCount; j++) {
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                if (StringUtils.equals(String.valueOf(cell.getCellType()), "STRING")) {
                                    if (StringUtils.equals(cell.getStringCellValue(), "プログラムID")) {
                                        row.getCell(j + 3).setCellValue(fileCode);
                                    }
                                    if (StringUtils.equals(cell.getStringCellValue(), "プログラム名")) {
                                        row.getCell(j + 3).setCellValue(fileName);
                                    }

                                }
                            }
                        }
                    }
                }
            }
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            exportFileExcel(workbook, path, fileOutputPathText);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, e.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void writingThemD5_2(Path path, String fileOutputPathText) {
        try (final Workbook workbook = new XSSFWorkbook(new FileInputStream(path.toFile()))) {
            final Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                final Sheet sheet = sheetIterator.next();
                // Lặp qua các dòng trong sheet
                for (int i = 0; i < sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        int cellCount = row.getPhysicalNumberOfCells();
                        for (int j = 0; j < cellCount; j++) {
                            Cell cell = row.getCell(j);
                            if (cell != null) {
                                if (StringUtils.equals(String.valueOf(cell.getCellType()), "STRING")) {
                                    if (StringUtils.equals(cell.getStringCellValue(), "【プログラム名称 or 共通プログラム名称】")) {
                                        row.getCell(j).setCellValue(fileName);
                                    }
                                    if (StringUtils.equals(cell.getStringCellValue(), "【プログラム記号名称 or 共通プログラム記号名】")) {
                                        row.getCell(j).setCellValue(fileCode);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
            exportFileExcel(workbook, path, fileOutputPathText);
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, e.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
        }
    }

    // export file excel
    public static void exportFileExcel(Workbook workbook, Path path, String fileOutputPathText) throws FileNotFoundException, IOException {
        String rsFileName = path.getFileName().toString().replace("【プログラム記号名称・共通プログラム記号名】", StringUtils.join(fileCode));
        Path filePathData = Paths.get(fileOutputPathText, rsFileName);
        try (FileOutputStream outputStream = new FileOutputStream(path.toFile(), false)) {
            workbook.write(outputStream);
        }
        Files.copy(path, filePathData, StandardCopyOption.REPLACE_EXISTING);
    }
}
