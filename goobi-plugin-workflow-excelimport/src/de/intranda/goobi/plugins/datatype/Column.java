package de.intranda.goobi.plugins.datatype;

public interface Column {
	public String getHeaderName();
    public String getRulesetName();
    public Integer getIdentifierColumn();


    public String getNormdataHeaderName();

    public String getDocType();
    
    
    public boolean isRequired();
    public String getPattern();
}
