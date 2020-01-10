import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Base64.Decoder;
import java.util.Date;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.json.JSONObject;

public class ReadExcel
{
  private static final Logger logger = Logger.getLogger("com.ibm.bpm.custom.ReadExcel");
  
  public ReadExcel() {}

  public String read(String base64ExcelData)
  {
    StringBuilder sb = new StringBuilder();
    
    if ((base64ExcelData == null) || (base64ExcelData.isEmpty()))
    {
      logger.logp(Level.OFF, "com.ibm.bpm.custom.ReadExcel", "read", "ReadExcel(read) - The Excel data passed is either missing or bad.");
      
      throw new RuntimeException("ReadExcel(read) - The Excel data passed is either missing or bad.");
    }
    

    byte[] data = java.util.Base64.getDecoder().decode(base64ExcelData);
    try {
      ByteArrayInputStream bais = new ByteArrayInputStream(data);
      

      Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(bais);
      
      JSONObject wbJSON = new JSONObject();
      JSONObject sheetJSON = null;
      JSONObject rowJSON = null;
      JSONObject cellJSON = null;
      
      List<JSONObject> sheetJSONList = new ArrayList();
      List<JSONObject> rowJSONList = new ArrayList();
      List<JSONObject> cellJSONList = new ArrayList();
      


      String type = null;
      
      DataFormatter dataFormatter = new DataFormatter();
      FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSSZ");
      Date cellDate = null;
      
      String cellValue = null;
      wbJSON.put("numsheets", workbook.getNumberOfSheets());
      
      for (Sheet sheet : workbook) {
        sheetJSON = new JSONObject();
        
        sheetJSON.put("name", sheet.getSheetName());
        sheetJSON.put("numrows", sheet.getPhysicalNumberOfRows());
        
        rowJSONList.clear();
        
        for (Row row : sheet)
        {
          rowJSON = new JSONObject();
          rowJSON.put("rownum", row.getRowNum());
          rowJSON.put("numcells", row.getPhysicalNumberOfCells());
          rowJSONList.add(rowJSON);
          

          cellJSONList.clear();
          


          for (Cell cell : row)
          {
            cellValue = dataFormatter.formatCellValue(cell);
            

            if (logger.isLoggable(Level.FINE))
            {
              sb.append("SHEETNAME ").append(sheet.getSheetName()).append("= ROWNUM =").append(row.getRowNum()).append(" ,CELL INFO TYPE=[").append(cell.getCellType());
              sb.append("] VALUE=[").append(dataFormatter.formatCellValue(cell));
              

              logger.logp(Level.FINE, "com.ibm.bpm.custom.ReadExcel", "read", sb.substring(0));
              sb.setLength(0);
            }
            
            switch (cell.getCellType())
            {
            case BOOLEAN: 
              if (DateUtil.isCellDateFormatted(cell))
              {
                cellDate = cell.getDateCellValue();
                cellValue = sdf.format(cellDate);
                type = "date";


              }
              else if (isInt(cellValue))
              {
                type = "integer";
              }
              else {
                type = "decimal";
              }
              
              break;
            
            case ERROR: 
              type = "string";
              break;
            
            case STRING: 
              cellValue = cellValue.toLowerCase();
              type = "boolean";
              break;
            
            case NUMERIC: 
              type = "null";
              break;
            
            case BLANK: 
              type = "null";
              break;
            
            case _NONE: 
              type = "error";
              break;
            

            case FORMULA: 
              cellValue = dataFormatter.formatCellValue(cell, evaluator);
              CellType fct = evaluator.evaluateFormulaCell(cell);
              
              if (logger.isLoggable(Level.FINE))
              {
                sb.append("FORMULA INFO, FORMULA=[").append(dataFormatter.formatCellValue(cell));
                sb.append("], FORMULA_VALUE=[").append(cellValue).append("], FORMULA_TYPE=[");
                sb.append(evaluator.evaluate(cell).getCellType().toString()).append("]");
                
                logger.logp(Level.FINE, "com.ibm.bpm.custom.ReadExcel", "read", sb.substring(0));
                sb.setLength(0);
              }
              

              switch (fct) {
              case ERROR: 
                type = "string";
                break;
              
              case STRING: 
                cellValue = cellValue.toLowerCase();
                type = "boolean";
                break;
              
              case BOOLEAN: 
                if (DateUtil.isCellDateFormatted(evaluator.evaluateInCell(cell))) {
                  cellDate = evaluator.evaluateInCell(cell).getDateCellValue();
                  cellValue = sdf.format(cellDate);
                  type = "date";


                }
                else if (isInt(cellValue))
                {
                  type = "integer";
                }
                else {
                  type = "decimal";
                }
                

                break;
              case _NONE: 
                type = "error";
                break;
              case FORMULA: case NUMERIC: 
              default: 
                type = "unknown";
              }
              
              
              break;
            
            default: 
              type = "unknown";
            }
            
            cellJSON = new JSONObject();
            cellJSON.put("colIndex", cell.getColumnIndex());
            cellJSON.put("value", cellValue);
            cellJSON.put("type", type);
            cellJSONList.add(cellJSON);
          }
          rowJSON.put("Cells", cellJSONList);
        }
        

        sheetJSON.put("Rows", rowJSONList);
        sheetJSONList.add(sheetJSON);
      }
      

      wbJSON.put("Sheets", sheetJSONList);
      workbook.close();
      
      if (logger.isLoggable(Level.FINE))
      {

        logger.logp(Level.FINER, "com.ibm.bpm.custom.ReadExcel", "read", wbJSON.toString());
      }
      

      return wbJSON.toString();
    }
    catch (IOException e)
    {
      logger.logp(Level.OFF, "com.ibm.bpm.custom.ReadExcel", "read", "error reading byte stream of Excel file", e);
    }
    




    return null;
  }
  

  private boolean isInt(String val)
  {
    boolean isInt = false;
    try
    {
      Integer.parseInt(val);
      isInt = true;
    }
    catch (NumberFormatException localNumberFormatException) {}
    



    return isInt;
  }
}