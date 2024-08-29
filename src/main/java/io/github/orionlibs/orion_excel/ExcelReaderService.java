package io.github.orionlibs.orion_excel;

import io.github.orionlibs.orion_object.ResourceCloser;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReaderService
{
    private XSSFWorkbook excelReader;
    private String sheetName;


    public ExcelReaderService()
    {
    }


    public ExcelReaderService(String fullFilePathAndName) throws IOException
    {
        this.excelReader = new XSSFWorkbook(fullFilePathAndName);
    }


    public ExcelReaderService(String fullFilePathAndName, String sheetName) throws IOException
    {
        this.excelReader = new XSSFWorkbook(fullFilePathAndName);
        this.sheetName = sheetName;
    }


    public ExcelReaderService(InputStream inputStream) throws IOException
    {
        this.excelReader = new XSSFWorkbook(inputStream);
    }


    public ExcelReaderService(InputStream inputStream, String sheetName) throws IOException
    {
        this.excelReader = new XSSFWorkbook(inputStream);
        this.sheetName = sheetName;
    }


    public void closeExcelFile()
    {
        ResourceCloser.closeResource(this.excelReader);
    }
    /*public List<String[]> getExcelRows()
    {
        Sheet sheet = excelReader.getSheetAt(0);
        Map<Integer, List<String>> data = new HashMap<>();
        //Row row = sheet.getRow(0);
        
        
        
        int i = 0;
        for (Row row : sheet) {
            data.put(i, new ArrayList<String>());
            for (Cell cell : row) {
                switch (cell.getCellType()) {
                case STRING:
                    data.get(i)
                        .add(cell.getRichStringCellValue()
                            .getString());
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        data.get(i)
                            .add(cell.getDateCellValue() + "");
                    } else {
                        data.get(i)
                            .add((int)cell.getNumericCellValue() + "");
                    }
                    break;
                case BOOLEAN:
                    data.get(i)
                        .add(cell.getBooleanCellValue() + "");
                    break;
                case FORMULA:
                    data.get(i)
                        .add(cell.getCellFormula() + "");
                    break;
                default:
                    data.get(i)
                        .add(" ");
                }
            }
            i++;
        }
        
        
        
        List<String[]> rows = excelReader.readAll();
        closeExcelFile();
        return rows;
    }*/


    public List<String[]> getExcelRowsExceptForHeader()
    {
        Sheet sheet = excelReader.getSheetAt(0);
        if(sheet == null)
        {
            sheet = excelReader.getSheet(sheetName);
        }
        List<String[]> rows = new ArrayList<>();
        Row headerRow = sheet.getRow(0);
        int numberOfColumns = headerRow.getPhysicalNumberOfCells();
        for(int i = 1; i < sheet.getPhysicalNumberOfRows(); i++)
        {
            Row row = sheet.getRow(i);
            String[] rowData = new String[numberOfColumns];
            for(int j = 0; j < numberOfColumns; j++)
            {
                if(row.getCell(j) != null)
                {
                    switch(row.getCell(j).getCellType())
                    {
                        case STRING:
                            rowData[j] = row.getCell(j).getStringCellValue();
                            break;
                        case NUMERIC:
                            if(DateUtil.isCellDateFormatted(row.getCell(j)))
                            {
                                rowData[j] = row.getCell(j).getDateCellValue() + "";
                            }
                            else
                            {
                                rowData[j] = row.getCell(j).getNumericCellValue() + "";
                            }
                            break;
                        case BOOLEAN:
                            rowData[j] = row.getCell(j).getBooleanCellValue() + "";
                            break;
                        case FORMULA:
                            rowData[j] = row.getCell(j).getCellFormula() + "";
                            break;
                        default:
                            rowData[j] = " ";
                    }
                }
                else
                {
                    rowData[j] = " ";
                }
            }
            rows.add(rowData);
        }
        closeExcelFile();
        return rows;
    }


    public List<String[]> getExcelHeadersRow()
    {
        Sheet sheet = excelReader.getSheetAt(0);
        if(sheet == null)
        {
            sheet = excelReader.getSheet(sheetName);
        }
        Row row = sheet.getRow(0);
        String[] rowData = new String[row.getPhysicalNumberOfCells()];
        for(int i = 0; i < row.getPhysicalNumberOfCells(); i++)
        {
            rowData[i] = row.getCell(i).getStringCellValue();
        }
        List<String[]> rows = new ArrayList<>();
        rows.add(rowData);
        closeExcelFile();
        return rows;
    }
    /*public List<String> getExcelColumn(int columnIndex)
    {
        List<String[]> rows = excelReader.readAll();
        closeExcelFile();
        List<String> column = new ArrayList<>();
        rows.forEach(row -> column.add(row[columnIndex]));
        return column;
    }*/
    /*public List<String> getExcelColumnExceptForHeader(int columnIndex)
    {
        List<String[]> rows = excelReader.readAll();
        closeExcelFile();
        List<String> column = new ArrayList<>();
        rows.forEach(row -> column.add(row[columnIndex]));
        return column.subList(1, rows.size());
    }*/
}