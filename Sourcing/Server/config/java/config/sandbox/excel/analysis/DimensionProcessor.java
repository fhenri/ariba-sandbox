package config.sandbox.excel.analysis;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import ariba.analytics.metadata.DimHierarchyMeta;
import ariba.analytics.metadata.DimLevelMeta;
import ariba.analytics.metadata.DimensionFieldMeta;
import ariba.analytics.metadata.DimensionTableMeta;
import ariba.analytics.metadata.FactDimFieldMeta;
import ariba.analytics.metadata.FactTableMeta;
import ariba.analytics.metadata.MetaDataConstants;
import ariba.analytics.metadata.VisibilityUtil;
import ariba.util.core.ListUtil;
import ariba.util.log.Log;

import config.sandbox.excel.ExcelHelper;

/**
 * Process dimension information. Read as much as possible of the dimension meta data and outputs
 * information in xl files. (look the Reporting Manager Fact/Dimension of Ariba Administration to
 * see information you can expect).
 * 
 * @author fhenri
 */
public class DimensionProcessor {
    
    private String dimName;
    private String sheetName;
    private HSSFWorkbook wkb;
    private HSSFSheet sheet;
    private ExcelHelper xlHelper;
    private DimensionTableMeta dim;
    private FactTableMeta[] facts;
    
    /**
     * Public constructor
     * 
     * @param name
     * @param dim
     * @param allFacts
     * @param xlHelper
     */
    public DimensionProcessor (String name, DimensionTableMeta dim, FactTableMeta[] allFacts, ExcelHelper xlHelper) {
        this.dimName = name;
        this.dim = dim;
        this.facts = allFacts;
        this.wkb = xlHelper.getWorkbook();
        this.xlHelper = xlHelper;
        initializeSheet();
    }

    // -----------------------------------------
    // GETTER METHODS
    // -----------------------------------------
    public String getDimensionName () {
        return dimName;
    }
    public String getDimDisplayName () {
        return dim.getDisplayName();
    }
    public HSSFSheet getDimensionSheet () {
        return sheet;
    }
    public String getSheetName () {
        return sheetName;
    }
    
    /**
     * Entry point of the class wich will read the dimension properties/field and write into xl.
     */
    public void process () {
        processProperties();
        processDimension();
    }
    
    /**
     * Initialize xl information
     */
    protected void initializeSheet () {
        Log.customer.debug("initialize sheet for dim %s", dimName);
        
        sheetName = "d." + dimName.substring(dimName.lastIndexOf('.')+1);
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
        
        row = sheet.createRow(5);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Hierarchies"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(6);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Hierarchy UniqueName"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Hierarchy Label"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString("Level UniqueName"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString("Level Label"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(5);
        cell.setCellValue(new HSSFRichTextString("Level Description"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        
        sheet.setColumnWidth(1, 50*256);
        sheet.setColumnWidth(2, 50*256);
        sheet.setColumnWidth(3, 40*256);
        sheet.setColumnWidth(4, 50*256);
        sheet.setColumnWidth(5, 50*256);
    }

    /**
     * write the header part of the dimension
     */
    private void processProperties () {
        
        HSSFRow row = sheet.getRow(1);
        row.createCell(2).setCellValue(new HSSFRichTextString(dim.getDisplayName()));
        row = sheet.getRow(2);
        row.createCell(2).setCellValue(new HSSFRichTextString(dimName));
    }
    
    /**
     * Process the dimension properties
     */
    private void processDimension () {
        
        // hierarchies
        DimHierarchyMeta[] hierarchies = dim.getDimHierarchies();
        int nbLinesOfHierarchies = writeHierarchies(hierarchies, 7);
        
        HSSFRow row = sheet.createRow(9+nbLinesOfHierarchies);
        HSSFCell cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Levels"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(10+nbLinesOfHierarchies);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Level UniqueName"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Field UniqueName"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString("Field Label"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(4);
        cell.setCellValue(new HSSFRichTextString("Field Type"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(5);
        cell.setCellValue(new HSSFRichTextString("Lookup Key"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        
        DimLevelMeta[] levelList = dim.getLevels();
        int nbLinesOfLevels = writeLevels(levelList, 11+nbLinesOfHierarchies);
        
        row = sheet.createRow(15+nbLinesOfHierarchies+nbLinesOfLevels);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Facts Information"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        row = sheet.createRow(16+nbLinesOfHierarchies+nbLinesOfLevels);
        cell = row.createCell(1);
        cell.setCellValue(new HSSFRichTextString("Fact Name"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(2);
        cell.setCellValue(new HSSFRichTextString("Fact UniqueName"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());
        cell = row.createCell(3);
        cell.setCellValue(new HSSFRichTextString("Fields Label"));
        cell.setCellStyle(xlHelper.getHeaderTableStyle());

        writeFactInformationForDimension(17+nbLinesOfHierarchies+nbLinesOfLevels);
    }
    
    /**
     * @param hierarchies
     * @param offset
     * @return
     */
    private int writeHierarchies (DimHierarchyMeta[] hierarchies, int offset)
    {
        HSSFRow row;
        HSSFCell cell;
        int levelOffset = 0;
        int validHierarchiesCount = 0;
        
        for (int h = 0; h < hierarchies.length; h++) {
            DimHierarchyMeta hierarchy = hierarchies[h];
            
            HSSFCellStyle style = xlHelper.getDefaultFieldStyle();
            if (!validHierarchy(hierarchy)) continue;
            validHierarchiesCount++;
            
            if (!hierarchy.isVisible()) {
                // set strike style
                style = xlHelper.getNonVisibleFieldStyle();
            }
                
            row = sheet.createRow(offset+h+levelOffset);
            //row = sheet.createRow(10*h);
            cell = row.createCell(0);
            String rank = 
                (String) hierarchy.getPropertyValue(MetaDataConstants.PropRank);
            cell.setCellValue(new HSSFRichTextString(rank));
            cell.setCellStyle(style);

            cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(hierarchy.getName()));
            cell.setCellStyle(style);
            cell = row.createCell(2);
            cell.setCellValue(new HSSFRichTextString(hierarchy.getDisplayName()));
            cell.setCellStyle(style);

            for (int l = 0; l < hierarchy.levels.length; l++) {
                DimLevelMeta level = hierarchy.levels[l];
                
                if (l!=0) {
                    row = sheet.createRow(offset+h+l+levelOffset);
                }
                //row = sheet.createRow(10*h+l);
                cell = row.createCell(3);
                cell.setCellValue(new HSSFRichTextString(level.getName()));
                cell.setCellStyle(style);
                cell = row.createCell(4);
                cell.setCellValue(new HSSFRichTextString(level.getDisplayName()));
                cell.setCellStyle(style);
                cell = row.createCell(5);
                cell.setCellValue(new HSSFRichTextString(level.getDescription()));
                cell.setCellStyle(style);
            }
            //sheet.groupRow(offset+h+levelOffset, hierarchy.levels.length);
            //sheet.setRowGroupCollapsed(offset+h+levelOffset, true);
            
            levelOffset += hierarchy.levels.length;
        }
        return levelOffset + validHierarchiesCount;
    }

    /**
	 * dimension's level properties
	 */
	private int writeLevels (DimLevelMeta[] levelList, int offset) {
        
        HSSFRow row;
        HSSFCell cell;
        int validLevelCount = 0;
        int levelOffset = 0;
        
        HSSFCellStyle style = xlHelper.getDefaultFieldStyle();
        for (int l = 0; l < levelList.length; l++) {
            DimLevelMeta level = levelList[l];
            if (!validLevel(level)) continue;
            
            validLevelCount++;
            row = sheet.createRow(offset+validLevelCount+levelOffset);
            cell = row.createCell(1);
            cell.setCellValue(new HSSFRichTextString(level.getName()));
            cell.setCellStyle(style);
            
            DimensionFieldMeta[] dimFields = (DimensionFieldMeta[])level.getFields();
            for (int f = 0; f < dimFields.length; f++) {
                DimensionFieldMeta field = dimFields[f];

                if (f!=0) {
                    row = sheet.createRow(offset+validLevelCount+f+levelOffset);
                }
                
                cell = row.createCell(2);
                cell.setCellValue(new HSSFRichTextString(field.getName()));
                cell.setCellStyle(style);
                cell = row.createCell(3);
                cell.setCellValue(new HSSFRichTextString(field.getDisplayName()));
                cell.setCellStyle(style);
                cell = row.createCell(4);
                cell.setCellValue(new HSSFRichTextString(field.getTypeName()));
                cell.setCellStyle(style);
                cell = row.createCell(5);
                if (level.isPartOfLookupKey(field)) {
                    cell.setCellValue(new HSSFRichTextString("true"));
                } else {
                    cell.setCellValue(new HSSFRichTextString("false"));
                }
                cell.setCellStyle(style);
            }
            levelOffset += dimFields.length;
        }
        return levelOffset + validLevelCount;
    }

    /**
     * @param offset
     */
    private void writeFactInformationForDimension(int offset) {
        int validFactCount=0;
        HSSFCellStyle style = xlHelper.getDefaultFieldStyle();
        
        for (int j = 0; j < facts.length; j++) {
            List<String> fields = ListUtil.list();
            boolean factHasField = false;
            FactTableMeta fact = facts[j];
            if (VisibilityUtil.isVisible(fact)) {
                for (int i = 0; i < fact.queryFields.length; i++) {
                    if (fact.queryFields[i] instanceof FactDimFieldMeta) {
                        FactDimFieldMeta dimRef = (FactDimFieldMeta)fact.queryFields[i];
                        if (dim == dimRef.getTableMeta()) {
                            factHasField = true;
                            fields.add(dimRef.getDisplayName());
                        }
                    }
                }
            }
            if (factHasField) {
                HSSFRow row = sheet.createRow(offset+(validFactCount++));
                HSSFCell cell = row.createCell(1);
                cell.setCellValue(new HSSFRichTextString(fact.getDisplayName()));
                cell.setCellStyle(style);
                cell = row.createCell(2);
                cell.setCellValue(new HSSFRichTextString(fact.getName()));
                cell.setCellStyle(style);
                cell = row.createCell(3);
                cell.setCellValue(new HSSFRichTextString(ListUtil.listToCSVString(fields)));
                cell.setCellStyle(style);
            }
        }
        
        
    }

    /**
     * validates a hierarchy. a hierarchy is considered invalid if one of its field is unused.
     * 
     * @param hierarchy
     * @return
     */
    private boolean validHierarchy (DimHierarchyMeta hierarchy) {
        for (int l = 0; l < hierarchy.levels.length; l++) {
            DimLevelMeta level = hierarchy.levels[l];
            if (!validLevel(level)) return false;
        }
        return true;
    }
    
    /**
     * validates a level. a level is considered invalid if one of its field is unused.
     * 
     * @param hierarchy
     * @return
     */
    private boolean validLevel (DimLevelMeta level) {
        DimensionFieldMeta[] dimFields = (DimensionFieldMeta[])level.getFields();
        for (int f = 0; f < dimFields.length; f++) {
            DimensionFieldMeta field = dimFields[f];
            if (field.isUnusedField()) return false;
        }
        return true;
    }
}
