package de.intranda.goobi.plugins.datatype;

import org.goobi.beans.User;

public class UserWrapper {
	private User user;
	private Boolean member;
	
	public UserWrapper(User inUser, Boolean inMember) {
		user = inUser;
		member = inMember;
	}
	
	public User getUser() {
		return user;
	}
	
	public void setUser(User user) {
		this.user = user;
	}
	
	public Boolean getMember() {
		return member;
	}
	
	public void setMember(Boolean member) {
		this.member = member;
	}
}
