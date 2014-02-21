/**
 * 
 */
package org.fhsolution.admin.excel;

import ariba.util.core.FileUtil;
import ariba.util.core.IOUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * convenient method to handle an xl report
 * 
 * @author fhenri
 *
 */
public class ExcelReport {
    
    protected HSSFWorkbook workbook;
    private String xlFileName;
    private File sourceFile;
    
    /**
     * Public constructor.
     * 
     * @param fileName name of the xl report
     */
    public ExcelReport (String fileName) {
    	this (fileName, fileName);
    }

    /**
     * Public constructor with a xl template
     * 
     * @param xlFileName
     * @param templateFileName
     */
    public ExcelReport (String xlFileName, String templateFileName) {
    	this.xlFileName = xlFileName;
    	sourceFile = new File(xlFileName);
    	
    	File templateFile = new File(templateFileName); 
        if (templateFile.exists()) {
            IOUtil.copyFile(templateFile, sourceFile);
            try {
                POIFSFileSystem poiFs = new POIFSFileSystem(IOUtil.bufferedInputStream(sourceFile));
                workbook = new HSSFWorkbook(poiFs);
            }
            catch (IOException e) {
                workbook = new HSSFWorkbook();
            }
        } else {
            workbook = new HSSFWorkbook();
        }
    }
    
    /**
     * save the report into xl file
     * 
     * @throws java.io.IOException
     */
    public void saveReport () throws IOException {
        //sourceFile = new File(xlFileName);
        FileUtil.createDirsForFile(sourceFile);
        FileOutputStream fileOut = new FileOutputStream(sourceFile);
        workbook.write(fileOut);
    }

    /**
     * Add an export report to the report
     * 
     * @param export
     */
    public void addExport (ExcelExport export) {
        export.initializeSheet(workbook);
        export.processExport();
    }
    
    /**
     * return the file for this xl report
     * 
     * @return
     */
    public File getFileReport () {
    	return sourceFile;
    }
    
    /**
     * Get the report name
     * 
     * @return
     */
    public String getReport () {
        return xlFileName;
    }

    @Override
    public String toString() {
        return "FileName[" + xlFileName + "]";
    }
}
