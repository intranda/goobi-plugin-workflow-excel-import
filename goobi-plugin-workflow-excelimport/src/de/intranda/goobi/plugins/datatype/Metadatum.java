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

	
	public String getErrorMessagesAsHtml() {
		String html="";
		if (errorMessages!=null) {
			html = "<ul class=\"popoverValidationList\">";
            for (String s : errorMessages) {
            	html += "<li>" + s + "</li>";
            }
            html += "</ul>";
		}
		return html;
	}
	
}
