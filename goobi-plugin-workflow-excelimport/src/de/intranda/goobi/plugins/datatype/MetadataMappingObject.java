package de.intranda.goobi.plugins.datatype;


import java.util.List;
import java.util.regex.Pattern;

import lombok.Data;

@Data
public class MetadataMappingObject {

    private String rulesetName;
    private String propertyName;
    private Integer excelColumn;

    private String headerName;
    private String identifier;
    private int columnNumber=-1;

    private String normdataHeaderName;

    private String docType ;
    
    
    private boolean required;
    private Pattern pattern;
    
    private List<String> validContent;
    private String eitherHeader;
    private String[] requiredHeaders;
    
    private String requiredErrorMessage;
    private String patternErrorMessage;
    private String validContentErrorMessage;
    private String eitherErrorMessage;
    private String requiredHeadersErrormessage;
    
}
