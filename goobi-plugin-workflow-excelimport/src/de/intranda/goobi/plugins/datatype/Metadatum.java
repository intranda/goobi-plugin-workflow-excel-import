package de.intranda.goobi.plugins.datatype;

import java.util.ArrayList;
import java.util.List;

import lombok.Data;

@Data
public class Metadatum {
	String headername;
	String value;
	boolean valid=true;
	List<String> errorMessages=new ArrayList<>();

}
