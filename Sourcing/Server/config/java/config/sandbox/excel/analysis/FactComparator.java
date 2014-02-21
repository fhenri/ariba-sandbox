package config.sandbox.excel.analysis;

import ariba.analytics.metadata.FactTableMeta;

/**
 * Comparables class for FactTableMeta object.
 * Fact are compared based on their names.
 * 
 * @author fhenri
 */
public class FactComparator implements Comparable<FactComparator> {

    private String factName;
    
    /**
     * Public constructor
     * 
     * @param fact
     */
    public FactComparator (FactTableMeta fact) {
        this.factName = fact.getDisplayName();
    }
    
    /**
     * fact name getter.
     * @return the name of the fact
     */
    public String getFactName () {
        return factName;
    }
    
    /**
     * compare a fact with another fact.
     * 
     * @see Comparable#compareTo(Object)
     */
    public int compareTo(FactComparator fc) {
        return this.factName.compareTo(fc.getFactName());
    }
}
