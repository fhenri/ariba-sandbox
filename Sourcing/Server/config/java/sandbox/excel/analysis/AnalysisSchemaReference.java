package sandbox.excel.analysis;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import ariba.analytics.metadata.DimensionTableMeta;
import ariba.analytics.metadata.FactTableMeta;
import ariba.analytics.metadata.MetaData;
import ariba.base.core.Base;
import ariba.user.core.User;
import ariba.util.core.FileUtil;
import ariba.util.core.MapUtil;
import ariba.util.core.StringUtil;
import ariba.util.log.Log;

import sandbox.excel.ExcelHelper;

/**
 * The Analysis Schema Reference gives the Analysis schema representation in xl file.
 * By default, it outputs all fact/dimension/fields as defined by the object model.
 * It is possible to restraint this output by
 *  - fact given factname
 *  - roles of a certain user, what a user with such role would be able to see
 *
 * @author fhenri
 */
public class AnalysisSchemaReference {

    private String fileName;
    
    private HSSFWorkbook workbook;
    private ExcelHelper xlHelper;
    
    private Map<FactComparator, FactProcessor> facts;
    private Map<String, DimensionProcessor> dimensions;
    
    private String userName;
    private String userPasswordAdapter;
    
    /**
     * Public Constructor with the filename
     * 
     * @param fileName
     */
    public AnalysisSchemaReference (String fileName) {
        this.fileName = fileName;
        facts = MapUtil.sortedMap();
        dimensions = MapUtil.sortedMap();
    }
    
    /**
     * Public Constructor with the filename and a list of fact to document
     * 
     * @param fileName
     * @param factList
     */
    public AnalysisSchemaReference (String fileName, List<String> factList) {
        this(fileName);
    }
    
    /**
     * Public Constructor with the filename and the user to use as reference
     * 
     * @param fileName
     * @param userName
     */
    public AnalysisSchemaReference (String fileName, String userName) {
        this(fileName);
        this.userName = userName;
        this.userPasswordAdapter = "PasswordAdapter1";
    }
    
    /**
     * Public Constructor with the filename, the fact and the user to use as reference
     * 
     * @param fileName
     * @param factList
     * @param userName
     */
    public AnalysisSchemaReference (
            String fileName, 
            List<String> factList, 
            String userName) {
        this(fileName, userName);
    }
    
    /**
     * Entry point of the program
     */
    @SuppressWarnings("unchecked")
    public void execute () {
        
        // create a session for user
        if (StringUtil.nullOrEmptyOrBlankString(userName)) {
            Log.customer.debug("Default User Session to aribasystem User");
            Base.getSession().setEffectiveUser(User.getAribaSystemUser().getBaseId());
        } else {
            User user = User.getUser(userName, userPasswordAdapter);
            Base.getSession().setEffectiveUser(user.getBaseId());
        }
        
        // build the config.sandbox.excel file
        workbook = new HSSFWorkbook();
        xlHelper = new ExcelHelper(workbook);

        // get all dimension tables
        Map<?,?> dimTables = MetaData.getMetaData().getAllDimensions();
        Iterator<?> dimKeyValuePairs = dimTables.entrySet().iterator();
        
        // get all fact tables 
        FactTableMeta[] allFacts = MetaData.getMetaData().factTableMetas;
        
        // loop on dimensions
        while (dimKeyValuePairs.hasNext()) {
            Map.Entry<String,DimensionTableMeta> aDim = 
                (Map.Entry<String, DimensionTableMeta>) dimKeyValuePairs.next();
            String dimName = aDim.getKey();
            DimensionTableMeta dimTable = aDim.getValue();
            DimensionProcessor processor = new DimensionProcessor (dimName, dimTable, allFacts, xlHelper);
            processor.process();
            dimensions.put(dimName, processor);
        }
        
        Map<String, FactTableMeta> factTables = MetaData.getMetaData().getAllFacts();
        Iterator<Map.Entry<String, FactTableMeta>> factKeyValuePairs = factTables.entrySet().iterator();
        // loop on fact
        while (factKeyValuePairs.hasNext()) {
            Map.Entry<String, FactTableMeta> aFact = factKeyValuePairs.next();
            String factName = aFact.getKey();
            FactTableMeta factTable = aFact.getValue();
            FactComparator factComp = new FactComparator(factTable);
            FactProcessor processor = new FactProcessor (factName, factTable, xlHelper, dimensions);
            processor.process();
            facts.put(factComp, processor);
        }

        createIndex();
        
        //save the config.sandbox.excel file
        try {
            File sourceFile = new File(fileName);
            FileUtil.createDirsForFile(sourceFile);
            FileOutputStream fileOut = new FileOutputStream(sourceFile);
            workbook.write(fileOut);
        } catch (IOException ioe) {
            Log.customer.error(10001);
        }

    }

    /**
     * create index with a link on each fact / dimension from the field
     * give name / description and link on the worksheet
     */
    private void createIndex() {
        
        // create the sheet for index
        HSSFSheet sheet = workbook.createSheet("Index");
        workbook.setSheetOrder("Index", 0);
        
        HSSFRow row;
        HSSFCell cell;
        
        sheet.setColumnWidth(1, 50*256);
        sheet.setColumnWidth(2, 60*256);
        sheet.setColumnWidth(4, 50*256);
        
        row = sheet.createRow(1);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Fact Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Fact System Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString("Permission"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());

        int i=2;
        
        Iterator<Map.Entry<FactComparator, FactProcessor>> factKeyName = facts.entrySet().iterator();
        while (factKeyName.hasNext()) {
            Map.Entry<FactComparator, FactProcessor> aFact = factKeyName.next();
            FactProcessor fProcessor = aFact.getValue();
            
            String factName = fProcessor.getFactName();
            String factSheetName = fProcessor.getSheetName();
            
            HSSFCellStyle style = xlHelper.getDefaultFieldStyle();
            if (!fProcessor.isFactVisible()) {
                style = xlHelper.getNonVisibleFieldStyle();
            }
            
            row = sheet.createRow(i++);
            cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(fProcessor.getFactDisplayName()));
            cell.setCellStyle(style);

            cell = row.createCell(2);
            cell.setCellFormula(xlHelper.createFormula(factSheetName, factName));
            cell.setCellStyle(style);
            cell.setCellStyle(xlHelper.getHyperLinkStyle());
            
            cell = row.createCell(4);
            cell.setCellValue(new HSSFRichTextString(fProcessor.getFactViewPermission()));
            cell.setCellStyle(style);

        }

        //write dimensions
        i++;
        row = sheet.createRow(i++);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Dimension Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Dimension System Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        
        Iterator<Map.Entry<String, DimensionProcessor>> dimKeyName = dimensions.entrySet().iterator();
        while (dimKeyName.hasNext()) {
            Map.Entry<String, DimensionProcessor> aFact = dimKeyName.next();
            DimensionProcessor dProcessor = aFact.getValue();
            
            String dimName = dProcessor.getDimensionName();
            String dimSheetName = dProcessor.getSheetName();
            
            row = sheet.createRow(i++);
            cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(dProcessor.getDimDisplayName()));

            cell = row.createCell(2);
            cell.setCellFormula(xlHelper.createFormula(dimSheetName, dimName));
            cell.setCellStyle(xlHelper.getHyperLinkStyle());
            
        }
    }
}
