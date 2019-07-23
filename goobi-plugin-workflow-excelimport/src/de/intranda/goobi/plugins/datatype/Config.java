package de.intranda.goobi.plugins.datatype;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Pattern;

import org.apache.commons.configuration.HierarchicalConfiguration;
import org.apache.commons.configuration.SubnodeConfiguration;

import lombok.Data;
import lombok.extern.log4j.Log4j;

@Data
@Log4j
public class Config {

    private String publicationType;
    private String collection;
    private int firstLine;
    private int identifierColumn;
    private int conditionalColumn;
    private int rowHeader;
    private int rowIdentifier;
    private int rowDataStart;
    private int rowDataEnd;
    private List<MetadataMappingObject> metadataList = new ArrayList<>();
    private String identifierHeaderName;

    private boolean useOpac = false;
    private String opacName;
    private String searchField;

    /**
     * loads the &lt;config&gt; block from xml file
     * 
     * @param xmlConfig
     */

    @SuppressWarnings("unchecked")
    public Config(SubnodeConfiguration xmlConfig) {

        publicationType = xmlConfig.getString("/publicationType", "Monograph");
        collection = xmlConfig.getString("/collection", "");
        firstLine = xmlConfig.getInt("/firstLine", 1);
        identifierHeaderName = xmlConfig.getString("/identifierHeaderName", null);
        rowIdentifier = xmlConfig.getInt("/rowIdentifier", 1);
        rowHeader = xmlConfig.getInt("/rowHeader", 1);
        rowDataStart = xmlConfig.getInt("/rowDataStart", 2);
        rowDataEnd = xmlConfig.getInt("/rowDataEnd", 20000);

        List<HierarchicalConfiguration> mml = xmlConfig.configurationsAt("//metadata");
        for (HierarchicalConfiguration md : mml) {
            metadataList.add(getMetadata(md));
        }
        useOpac = xmlConfig.getBoolean("/useOpac", false);
        if (useOpac) {
            opacName = xmlConfig.getString("/opacName", "ALMA WUW");
            searchField = xmlConfig.getString("/searchField", "12");
        }
    }

    private MetadataMappingObject getMetadata(HierarchicalConfiguration md) {
        String rulesetName = md.getString("@ugh");
        String propertyName = md.getString("@name");
        Integer columnNumber = md.getInteger("@column", null);
        String headerName = md.getString("@headerName", null);
        String normdataHeaderName = md.getString("@normdataHeaderName", null);
        String docType = md.getString("@docType", "child");
        boolean required = md.getBoolean("@required", false);
        String patternString = md.getString("@pattern", "");
        String eitherHeader = md.getString("@either", "");
        String requiredFields = md.getString("@requiredFields", "");
        String listPath = md.getString("@list");
        String identifier = md.getString("@identifier");
        Integer wordcount = md.getInteger("@wordcount", 0);
        boolean split=md.getBoolean("@split",false);

        String requiredErrorMessage = md.getString("@requiredErrorMessage", "");
        String patternErrorMessage = md.getString("@patternErrorMessage", "");
        String validContentErrorMessage = md.getString("@listErrorMessage", "");
        String eitherErrorMessage = md.getString("@eitherErrorMessage", "");
        String requiredHeadersErrormessage = md.getString("@requiredFieldsErrormessage", "");
        String wordcountErrormessage = md.getString("@wordcountErrorMessage", "");
        ArrayList<String> validContent = new ArrayList<>();

        if (listPath != null && !listPath.isEmpty()) {
            try {
                validContent = readFileToList(listPath);
            } catch (FileNotFoundException e) {
                log.error("Configured list of options not found: "+listPath,e);
            }
        }

        String[] requiredHeaders = null;
        if (requiredFields != null) {
            requiredHeaders = requiredFields.split("; ");
        }
        Pattern pattern = null;
        if (!patternString.isEmpty()) {
            pattern = Pattern.compile(patternString);
        }

        MetadataMappingObject mmo = new MetadataMappingObject();
        mmo.setExcelColumn(columnNumber);
        mmo.setPropertyName(propertyName);
        mmo.setRulesetName(rulesetName);
        mmo.setHeaderName(headerName);
        mmo.setIdentifier(identifier);
        mmo.setNormdataHeaderName(normdataHeaderName);
        mmo.setDocType(docType);
        mmo.setRequired(required);
        mmo.setPattern(pattern);
        mmo.setValidContent(validContent);
        mmo.setEitherHeader(eitherHeader);
        mmo.setWordcount(wordcount);
        mmo.setSplit(split);
        mmo.setEitherErrorMessage(eitherErrorMessage);
        mmo.setPatternErrorMessage(patternErrorMessage);
        mmo.setRequiredErrorMessage(requiredErrorMessage);
        mmo.setRequiredHeadersErrormessage(requiredHeadersErrormessage);
        mmo.setValidContentErrorMessage(validContentErrorMessage);
        mmo.setWordcountErrormessage(wordcountErrormessage);
        if (requiredHeaders != null) {
            mmo.setRequiredHeaders(requiredHeaders);
        }
        return mmo;
    }

    private ArrayList<String> readFileToList(String listPath) throws FileNotFoundException {
        Scanner s = new Scanner(new File(listPath));
        ArrayList<String> validContent = new ArrayList<>();
        while (s.hasNext()) {
            validContent.add(s.next());
        }
        s.close();
        return validContent;
    }



}
