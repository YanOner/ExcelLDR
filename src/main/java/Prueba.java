import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.util.stream.Stream;

public class Prueba {

    public static void main(String[] args) {

        String pathFiles = "D:\\excels";

        try (Stream<Path> paths = Files.walk(Paths.get(pathFiles))) {
            paths
                    .filter(Files::isRegularFile)
                    .forEach(path -> {
                        leerExcel(pathFiles, path.getFileName().toString());
                    });
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static void leerExcel(String pathFile, String fileName) {

        try (FileInputStream file = new FileInputStream(new File(pathFile+"//"+fileName))) {
            // leer archivo excel
            XSSFWorkbook worbook = new XSSFWorkbook(file);
            //obtener la hoja que se va leer
            XSSFSheet sheet = worbook.getSheetAt(0);
            //obtener todas las filas de la hoja excel

            XSSFRow row;
            //se recorre cada fila hasta el final
            String lastHashValue = "";
            int count = 0;
            // Archivo nuevo
            XSSFWorkbook newWorbook = new XSSFWorkbook();
            XSSFSheet newSheet = newWorbook.createSheet("anexo5");
            int rowSize = sheet.getLastRowNum();
            int rowOld = 0;
            for(int i = 0;i<rowSize;i++) {
                row = sheet.getRow(i);
                XSSFRow newRow = newSheet.createRow(i);

                Cell cell = row.getCell(2);
                String cellHashValue = cell.getStringCellValue();

                if (cellHashValue.equals("q3lcvww7")) {
                    int x = 0;
                }

                copyCells(row, newRow);

                if (lastHashValue.equals(cellHashValue)) {
//                    System.out.println("cellHashValue: " + cellHashValue);
                    //Motivo
                    Cell cellMotivo = row.getCell(4);
                    Cell cellNuevoMotivo = newSheet.getRow(rowOld).getCell(4);
                    String cellMotivoValue = cellMotivo==null?"":cellMotivo.getStringCellValue();
                    String cellNuevoMotivoValue = cellNuevoMotivo.getStringCellValue();
                    if (cellNuevoMotivoValue.indexOf(cellMotivoValue) == -1) {
                        if (StringUtils.isBlank(cellMotivoValue)) {
                            cellNuevoMotivo.setCellValue(cellMotivoValue);
                        } else {
                            if (StringUtils.isBlank(cellNuevoMotivoValue)) {
                                cellNuevoMotivo.setCellValue(cellMotivoValue);
                            } else {
                                cellNuevoMotivo.setCellValue(cellNuevoMotivoValue + ", " + cellMotivoValue);
                            }
                        }
                    }
                    //Otros
                    Cell cellOtros = row.getCell(5);
                    String cellOtrosValue = cellOtros==null?"":cellOtros.getStringCellValue();
                    if (!StringUtils.isBlank(cellOtrosValue) && cellNuevoMotivoValue.indexOf("otros") == -1) {
                        if (StringUtils.isBlank(cellMotivoValue)) {
                            cellNuevoMotivo.setCellValue("otros");
                        } else {
                            cellNuevoMotivo.setCellValue(cellNuevoMotivoValue + ", otros");
                        }
                    }

                    //Unidad Organica
                    Cell cellUO = row.getCell(13);
                    Cell cellNuevoUO = newSheet.getRow(rowOld).getCell(13);
                    String cellUOValue;
                    if (cellUO != null && cellUO.getCellType().equals(CellType.NUMERIC)) {
                        cellUOValue = String.valueOf(cellUO.getNumericCellValue());
                    } else {
                        cellUOValue = cellUO==null?"":cellUO.getStringCellValue();
                    }

                    String cellNuevoUOValue;
                    if (cellNuevoUO != null && cellNuevoUO.getCellType().equals(CellType.NUMERIC)) {
                        cellNuevoUOValue = String.valueOf(cellNuevoUO.getNumericCellValue());
                    } else {
                        cellNuevoUOValue = cellNuevoUO==null?"":cellNuevoUO.getStringCellValue();
                    }

                    if (!cellNuevoUOValue.contains(cellUOValue)) {
                        if (StringUtils.isBlank(cellUOValue)) {
                            cellNuevoUO.setCellValue(cellUOValue);
                        } else {
                            if (StringUtils.isBlank(cellNuevoUOValue)) {
                                cellNuevoUO.setCellValue(cellUOValue);
                            } else {
                                cellNuevoUO.setCellValue(cellNuevoUOValue + ", " + cellUOValue);
                            }
                        }
                    }

                    //Tipo Respuesta
                    Cell cellTR = row.getCell(12);
                    Cell cellNuevoTR = newSheet.getRow(rowOld).getCell(12);
                    String cellTRValue;
                    if (cellTR != null && cellTR.getCellType().equals(CellType.NUMERIC)) {
                        cellTRValue = String.valueOf(cellTR.getNumericCellValue());
                    } else {
                        cellTRValue = cellTR==null?"":cellTR.getStringCellValue();
                    }

                    String cellNuevoTRValue;
                    if (cellNuevoTR != null && cellNuevoTR.getCellType().equals(CellType.NUMERIC)) {
                        cellNuevoTRValue = String.valueOf(cellNuevoTR.getNumericCellValue());
                    } else {
                        cellNuevoTRValue = cellNuevoTR==null?"":cellNuevoTR.getStringCellValue();
                    }

                    if (!cellNuevoTRValue.contains(cellTRValue)) {
                        if (StringUtils.isBlank(cellTRValue)) {
                            cellNuevoTR.setCellValue(cellTRValue);
                        } else {
                            if (StringUtils.isBlank(cellNuevoTRValue)) {
                                cellNuevoTR.setCellValue(cellTRValue);
                            } else {
                                cellNuevoTR.setCellValue(cellNuevoTRValue + ", " + cellTRValue);
                            }
                        }
                    }

                    if (rowOld == 0) {
                        rowOld = i;//fila actual
                    }
                    //borrar fila nueva
                    newSheet.removeRow(newRow);
                    count++;
                } else {
                    lastHashValue = cellHashValue;
                    rowOld = i;
                    count = 0;
                    //motivos otros
                    Cell cellMotivos1 = newRow.getCell(4);
                    Cell cellOtros1 = newRow.getCell(5);
                    String cellMotivos1Value = cellMotivos1.getStringCellValue();
                    String cellOtros1Value = cellOtros1.getStringCellValue();
                    if (StringUtils.isBlank(cellMotivos1Value) && !StringUtils.isBlank(cellOtros1Value)) {
                        cellMotivos1.setCellValue("otros");
                    }

                }
            }
//            System.out.println("Repetidos: " + count);

            File newFile = new File(pathFile+"//new-"+fileName);
            try (FileOutputStream fileOuS = new FileOutputStream(newFile)){
                if (newFile.exists()) {// si el archivo existe se elimina
                    newFile.delete();
                }
                newWorbook.write(fileOuS);
                fileOuS.flush();
                fileOuS.close();
                System.out.println(fileName + " Archivo Creado " + LocalDateTime.now());
            } catch (Exception e) {
                throw e;
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void copyCells(XSSFRow row, XSSFRow newRow) {
        int countCell = 0;
        for(int i = 0;i<20;i++) {
            XSSFCell cell = row.getCell(i);
            Cell newCell = newRow.createCell(countCell);
            if (cell != null && cell.getCellType().equals(CellType.NUMERIC)) {
                if (i == 9) {//cell dni
                    String isDNI = newRow.getCell(8).getStringCellValue();
                    if (isDNI.equals("DNI")) {
                        try {
                            int cd = (int) cell.getNumericCellValue();
                            String pl = String.format("%08d", cd);
                            newCell.setCellValue(pl);
                        } catch (Exception e) {
                            newCell.setCellValue(cell.getNumericCellValue());
                        }
                    } else {
                        try {
                            int cd = (int) cell.getNumericCellValue();
                            String pl = String.format("%09d", cd);
                            newCell.setCellValue(pl);
                        } catch (Exception e) {
                            newCell.setCellValue(cell.getNumericCellValue());
                        }                    }
                } else {
                    newCell.setCellValue(cell.getNumericCellValue());
                }
            } else {
                if (cell == null || StringUtils.isBlank(cell.getStringCellValue())) {
                    newCell.setCellValue("");
                } else {
                    newCell.setCellValue(cell.getStringCellValue());
                }
            }
            countCell++;
        }
    }

}
