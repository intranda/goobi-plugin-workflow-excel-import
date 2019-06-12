package de.intranda.goobi.plugins.datatype;


import java.util.List;

import lombok.Data;

@Data
public class MetadataMappingObject implements Column{

    private String rulesetName;
    private String propertyName;
    private Integer excelColumn;
    private Integer identifierColumn;

    private String headerName;

    private String normdataHeaderName;

    private String docType ;
    
    
    private boolean required;
    private String pattern;
    
    private List<String> validContent;
    private String eitherHeader;
    private String[] requiredHeaders;
}
