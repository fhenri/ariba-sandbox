/**
 *
 */
package org.fhsolution.admin.excel.user;

import ariba.base.core.Base;
import ariba.base.core.BaseId;
import ariba.base.core.Partition;
import ariba.base.core.aql.AQLOptions;
import ariba.base.core.aql.AQLQuery;
import ariba.base.core.aql.AQLResultCollection;
import ariba.user.core.User;
import ariba.util.core.Date;
import ariba.util.core.StringUtil;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;

import java.util.List;
import java.util.Locale;

/**
 *
 * @author fhenri
 *
 */
public class SupplierOrganizationExport extends OrganizationExport {

    /**
     * @param sheetName
     */
    public SupplierOrganizationExport(String sheetName) {
        super(sheetName);

        // this should be read from the metadata
        additionalFields =
                "";

        whereCriteria =
                " where o.IsSupplier = true";
    }

    protected void writeResultInRow (HSSFRow row, AQLResultCollection result) {
        super.writeResultInRow(row, result);

    }

    /**
     * @param row
     * @param subQuery
     * @param cell
     */
    protected void writeSubQueryResultInRow (HSSFRow row, String subQuery, int cell) {

        AQLQuery query = AQLQuery.parseQuery(subQuery);
        AQLOptions options = new AQLOptions(Partition.None);
        options.setPartition(Partition.Any);
        options.setUserLocale(Locale.ENGLISH);
        options.setUserPartition(Partition.Any);
        AQLResultCollection results = Base.getService().executeQuery(query, options);

        // loop on results
        String resultCellValue = "";
        while (results.next()) {
            String result = results.getString(0);
            if (!StringUtil.nullOrEmptyOrBlankString(result)) {
                resultCellValue = resultCellValue.concat(result).concat("\n");
            }
        }

        row.createCell(cell).setCellValue(new HSSFRichTextString(resultCellValue));

    }

    private HSSFRichTextString getStringDate (Object fieldValue) {
        if (fieldValue instanceof Date && fieldValue != null) {
            return new HSSFRichTextString(((Date)fieldValue).toString());
        }
        return null;
    }

    private String getUserString (List<BaseId> users) {
        String userString = "";
        for (BaseId baseId : users) {
            User leadBuyer = (User)Base.getSession().objectFromId(baseId);
            userString = userString.concat(leadBuyer.getName().getPrimaryString()).concat("\n");
        }
        return userString;
    }
}
