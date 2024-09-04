package io.github.orionlibs.orion_excel;

import io.github.orionlibs.orion_object.ResourceCloser;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriterService
{
    private XSSFWorkbook excelReader;
    private String sheetName;


    public ExcelWriterService()
    {
    }


    public ExcelWriterService(String fullFilePathAndName) throws IOException
    {
        this.excelReader = new XSSFWorkbook(fullFilePathAndName);
    }


    public ExcelWriterService(String fullFilePathAndName, String sheetName) throws IOException
    {
        this.excelReader = new XSSFWorkbook(fullFilePathAndName);
        this.sheetName = sheetName;
    }


    public void closeExcelFile()
    {
        ResourceCloser.closeResource(this.excelReader);
    }


    public void saveToExcelFile(List<Object[]> rows, String fullFilePathAndName, String sheetName) throws IOException
    {
        FileOutputStream out = new FileOutputStream(fullFilePathAndName);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet worksheet = workbook.createSheet(sheetName);
        int rowNum = 0;
        for(Object[] rowData : rows)
        {
            XSSFRow rowTemp = worksheet.createRow(rowNum);
            for(int j = 0; j < rowData.length; j++)
            {
                if(rowData[j] instanceof String data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else if(rowData[j] instanceof Byte data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else if(rowData[j] instanceof Short data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else if(rowData[j] instanceof Integer data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else if(rowData[j] instanceof Long data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else if(rowData[j] instanceof Float data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else if(rowData[j] instanceof Double data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else if(rowData[j] instanceof Boolean data)
                {
                    rowTemp.createCell(j).setCellValue(data);
                }
                else
                {
                    rowTemp.createCell(j).setCellValue(rowData[j].toString());
                }
            }
            ++rowNum;
        }
        workbook.write(out);
        out.flush();
        out.close();
    }
}