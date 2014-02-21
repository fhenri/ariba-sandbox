package sandbox.excel.analysis;

import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.SortedSet;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import ariba.analytics.metadata.FactDimFieldMeta;
import ariba.analytics.metadata.FactFieldMeta;
import ariba.analytics.metadata.FactFieldOrFormula;
import ariba.analytics.metadata.FactInlineDimFieldMeta;
import ariba.analytics.metadata.FactMeasureFieldMeta;
import ariba.analytics.metadata.FactMeasureFormulaMeta;
import ariba.analytics.metadata.FactTableMeta;
import ariba.analytics.metadata.MetaDataConstants;
import ariba.analytics.metadata.VisibilityUtil;
import ariba.util.core.SetUtil;
import ariba.util.log.Log;

import sandbox.excel.ExcelHelper;

/**
 * Fact sheet will contain:
 * - fact name
 * - class name 
 * - associated permission
 * - List of measure fields
 *   system name | name | type | comment (description by default)
 * - List of measure formula
 *   system name | name | formula | comment
 * - List of dimensions
 *   system name | name | type (link if any) | commment (description by default)
 * - List of other fields
 * 
 * @author fhenri
 */
public class FactProcessor {

    private String factName;
    private String sheetName;
    private String viewPermission;
    private HSSFWorkbook wkb;
    private HSSFSheet sheet;
    private ExcelHelper xlHelper;
    private FactTableMeta fact;
    private Map<String, DimensionProcessor> dimensions;
    
    /**
     * Public constructor
     * 
     * @param name
     * @param fact
     * @param xlHelper
     * @param dimensions
     */
    public FactProcessor (
            String name, 
            FactTableMeta fact, 
            ExcelHelper xlHelper, 
            Map<String, DimensionProcessor> dimensions) 
    {
        this.factName = name;
        this.fact = fact;
        this.wkb = xlHelper.getWorkbook();
        this.xlHelper = xlHelper;
        this.dimensions = dimensions;
        initializeSheet();
    }
    
    // -----------------------------------------
    // GETTER METHODS
    // -----------------------------------------
    public String getFactName () {
        return factName;
    }
    public String getFactDisplayName () {
        return fact.getDisplayName();
    }
    public String getFactDescription () {
        return fact.getDescription();
    }
    public boolean isFactVisible () {
        return VisibilityUtil.isVisible(fact);
    }
    public String getSheetName () {
        return sheetName;
    }
    public HSSFSheet getSheet () {
        return sheet;
    }
    public String getFactViewPermission () {
        return viewPermission;
    }
    
    /**
     * Entry point of the class wich will read the fact properties/field and write into xl.
     */
    public void process () {
        processProperties();
        processFact();
    }
    
    /**
     * Initialize the xl sheet and all the header part
     */
    private void initializeSheet () {
        Log.customer.debug("initialize sheet for fact %s", factName);
        
        sheetName = factName.substring(factName.lastIndexOf('.')+1);
        sheet = wkb.createSheet(sheetName);
        
        HSSFRow row;
        HSSFCell cell;
        
        row = sheet.createRow(1);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(2);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("ClassName"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(3);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Permission"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        
        row = sheet.createRow(5);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Measure Fields"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(6);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("System Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString("Type"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString("Comment"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        
        row = sheet.createRow(8);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Measure Formula"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(9);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("System Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString("Formula"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString("Comment"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());

        row = sheet.createRow(11);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Dimension Fields"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(12);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("System Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString("Type"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString("Comment"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());

        row = sheet.createRow(14);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Other Fields"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(15);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("System Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString("Type"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString("Comment"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());

        sheet.setColumnWidth(1, 50*256);
        sheet.setColumnWidth(2, 50*256);
        sheet.setColumnWidth(3, 40*256);
        sheet.setColumnWidth(4, 50*256);
    }

    /**
     * Process the main properties of the fact
     */
    private void processProperties () {
        
        // read  fact UI Name
        HSSFRow row = sheet.getRow(1);
        row.createCell(2).setCellValue(new HSSFRichTextString(fact.getDisplayName()));
        // read fact ClassName
        row = sheet.getRow(2);
        row.createCell(2).setCellValue(new HSSFRichTextString(factName));
        // read associate permission
        viewPermission = 
            (String) fact.getPropertyValue(MetaDataConstants.PropAnalysisViewPermission);
        row = sheet.getRow(3);
        row.createCell(2).setCellValue(new HSSFRichTextString(viewPermission));
    }
    
    /**
     * Process the fields from the fact
     */
    @SuppressWarnings("unchecked")
    private void processFact () {
        
        SortedSet<FieldComparator> otherFieldList = SetUtil.sortedSet();
        SortedSet<FieldComparator> queryFieldList = SetUtil.sortedSet();
        SortedSet<FieldComparator> measureList = SetUtil.sortedSet();
        SortedSet<FieldComparator> dimensionList = SetUtil.sortedSet();

        Iterator fields = fact.getAllFields(false);
        while (fields.hasNext()) {
            FactFieldMeta field = (FactFieldMeta)fields.next();
            FieldComparator fc = new FieldComparator(field);
            if (field instanceof FactDimFieldMeta) {
                dimensionList.add(fc);
            }
            else if (field instanceof FactMeasureFieldMeta) {
                measureList.add(fc);
            }
            else if (field instanceof FactInlineDimFieldMeta) {
                if (((FactInlineDimFieldMeta) field).isDeclaredDimension()) {
                    queryFieldList.add(fc);
                } else { 
                    otherFieldList.add(fc);
                }
            }
        }

        // read MeasureFields
        int measureField = writeField(measureList, 7, true);
        
        // read Measure Formulas
        List<FactFieldOrFormula> allFormulas = fact.getMeasureFormulas();
        int measureFormula = writeFormula(allFormulas, 10+measureField);
        
        // read DimensionReferences
        dimensionList.addAll(queryFieldList);
        int dimensionField = writeField(dimensionList, 13+measureField+measureFormula, true);
        
        // read OtherFields
        writeField(otherFieldList, 16+measureField+measureFormula+dimensionField, false);
    }
    
    /**
     * write field properties in xl
     * 
     * @param fieldList
     * @param offset
     * @param shift
     * @return
     */
    private int writeField (SortedSet<FieldComparator> fieldList, int offset, boolean shift) {
        HSSFRow row;
        HSSFCell cell;
        int validFieldCount = 0;
        
        Iterator<FieldComparator> fieldListIterator = fieldList.iterator();
        while (fieldListIterator.hasNext()) {
            FieldComparator fc = (FieldComparator) fieldListIterator.next();
            FactFieldMeta field = fc.getField();
            
            HSSFCellStyle style = xlHelper.getDefaultFieldStyle();
            if (field.isUnusedField()) continue;
            
            if (!field.isVisible()) {
                // set strike style
                style = xlHelper.getNonVisibleFieldStyle();
            }
            
            // we introduced the shifting for lines that needs to be inserted
            // when writing other field which are at the end of the file, it can get out of range
            // if we try to shift
            if (shift) {
                sheet.shiftRows(offset+validFieldCount, sheet.getLastRowNum(), 1);
            }
                
            row = sheet.createRow(offset+(validFieldCount++));
            cell = row.createCell(0);
            String rank = 
                (String) field.getPropertyValue(MetaDataConstants.PropRank);
            cell.setCellValue(new HSSFRichTextString(rank));
            cell.setCellStyle(style);
            
            cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(field.getName()));
            cell.setCellStyle(style);
            cell = row.createCell(2);
            cell.setCellValue(new HSSFRichTextString(field.getDisplayName()));
            cell.setCellStyle(style);
            
            String typeName = "";
            if (field instanceof FactDimFieldMeta) {
                FactDimFieldMeta fdfm = (FactDimFieldMeta) field;
                typeName = fdfm.getTableMeta().getName();
            } else {
                typeName = field.getTypeName();
            }            
            cell = row.createCell(3);
            cell.setCellValue(new HSSFRichTextString(typeName));
            cell.setCellStyle(style);
            if (dimensions.containsKey(typeName)) {
                DimensionProcessor dimProcessor = (DimensionProcessor) dimensions.get(typeName);
                String targetSheetName = dimProcessor.getSheetName();
                cell.setCellFormula(xlHelper.createFormula(targetSheetName, typeName));
                cell.setCellStyle(xlHelper.getHyperLinkStyle());
            }

            cell = row.createCell(4);
            cell.setCellValue(new HSSFRichTextString(field.getDescription()));
            cell.setCellStyle(style);

        }
        return validFieldCount;
    }
    
    /**
     * write formula properties in xl
     * 
     * @param formulaList
     * @param offset
     * @return
     */
    private int writeFormula (List<FactFieldOrFormula> formulaList, int offset) {
        HSSFRow row;
        int validFormulaCount = 0;
        
        for (int i=0; i<formulaList.size(); i++) {
            FactMeasureFormulaMeta fm = (FactMeasureFormulaMeta) formulaList.get(i);

            HSSFCellStyle style = xlHelper.getDefaultFieldStyle();
            if (formulaAvailable(fm) == 2) {
                continue;
            }
            validFormulaCount++;
            if (formulaAvailable(fm) == 1) {
                style = xlHelper.getNonVisibleFieldStyle();
            }
            sheet.shiftRows(i+offset, sheet.getLastRowNum(), 1);
            row = sheet.createRow(i+offset);
            HSSFCell cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(fm.getName()));
            cell.setCellStyle(style);
            cell = row.createCell(2);
            cell.setCellValue(new HSSFRichTextString(fm.getDisplayName()));
            cell.setCellStyle(style);
            cell = row.createCell(3);
            cell.setCellValue(new HSSFRichTextString(fm.getFormula()));
            cell.setCellStyle(style);
            cell = row.createCell(4);
            cell.setCellValue(new HSSFRichTextString(fm.getDescription()));
            cell.setCellStyle(style);
        }
        return validFormulaCount;
    }
    
    /**
     * check if the formula is available given the constraints and return the following:
     * 0 if the formula is available
     * 1 if the formula contains an invisibe field for the user
     * 2 if the formula contains an unused field
     * 
     * @param formula
     * @return
     */
    private int formulaAvailable (FactMeasureFormulaMeta formula) {
        FactMeasureFieldMeta[] fields = formula.getFields();
        for (FactMeasureFieldMeta field:fields) {
            if (field.isUnusedField()) return 2;
            if (!field.isVisible()) return 1;
        }
        return 0;
    }
}
