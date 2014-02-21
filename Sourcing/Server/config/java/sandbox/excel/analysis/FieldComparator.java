package sandbox.excel.analysis;

import ariba.analytics.metadata.FactFieldMeta;
import ariba.analytics.metadata.MetaDataConstants;
import ariba.util.core.StringUtil;

/**
 * Comparables field for FactFieldMeta object.
 * Fields are compared based on their ranking properties.
 * 
 * @author fhenri
 */
public class FieldComparator implements Comparable<FieldComparator> {

    private FactFieldMeta field;
    private int fieldRank;
    
    /**
     * Public constructor
     * 
     * @param field
     */
    public FieldComparator (FactFieldMeta field) {
        this.field = field;
        
        String rank = 
            (String) field.getPropertyValue(MetaDataConstants.PropRank);
        if (StringUtil.nullOrEmptyOrBlankString(rank)) {
            fieldRank = -1;
        } else {
            fieldRank = Integer.parseInt(rank);
        }
    }
    
    // -----------------------------------------
    // GETTER METHODS
    // -----------------------------------------
    public FactFieldMeta getField () {
        return field;
    }
    public int getRank () {
        return fieldRank;
    }
    
    /**
     * Compare a fact field with another fact field based on the rank property.
     * 
     * @see Comparable#compareTo(Object)
     */
    public int compareTo(FieldComparator fc) {
        if (this.fieldRank < fc.getRank()) return -1; 
        else return 1;
    }
    
    /**
     * @see Object#toString()
     */
    public String toString () {
        return field.toString();
    }
}
