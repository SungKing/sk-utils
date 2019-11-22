package cn.org.wangsong;

import cn.org.wangsong.func.DealCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.Objects;

/**
 * @Author Created by song.wang
 * @Create Date 2019-11-21 16:22
 */
public class ExcelUtil {

    /**
     * 处理 .xlsx
     * @param file
     * @param dealCell
     * @throws IOException
     * @throws InvalidFormatException
     */
    public void analyzeExcel(File file, DealCell dealCell) throws IOException, InvalidFormatException {

        XSSFWorkbook book = new XSSFWorkbook(file);
        int numberOfSheets = book.getNumberOfSheets();
        XSSFSheet sheet = book.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();
        XSSFRow row = sheet.getRow(0);
        short lastCellNum = row.getLastCellNum();
        for(int i =0;i<=lastRowNum;i++){
            XSSFRow row1 = sheet.getRow(i);
            if (row1!=null){
                for(int j = 0;j<=lastCellNum;j++){
                    XSSFCell cell = row1.getCell(j);
                    if (cell!=null){
                        dealCell.apply(i,j,getCellFormatValue(cell));
                    }
                }
            }

        }
    }

    /**
     *
     * @param cell 传入的列
     * @return 都是string类型
     */
    private String getCellFormatValue(XSSFCell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            if(Objects.equals(cell.getCellType(), CellType.STRING)){
                cellValue = cell.getRichStringCellValue().getString();
            }else if (Objects.equals(cell.getCellType(),CellType.NUMERIC)){
                cellValue = String.valueOf(cell.getNumericCellValue());
            }else if(Objects.equals(cell.getCellType(),CellType.FORMULA)){
                //判断cell是否为日期格式
                if(DateUtil.isCellDateFormatted(cell)){
                    //转换为日期格式YYYY-mm-dd
                    cellValue = cell.getDateCellValue();
                }else{
                    //数字
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
            }else{
                cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue.toString();
    }
}
