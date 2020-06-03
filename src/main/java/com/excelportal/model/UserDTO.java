package com.excelportal.model;

import javax.validation.constraints.NotEmpty;

import com.excelportal.annotations.FieldMatch;
import com.excelportal.annotations.ValidPassword;


@FieldMatch.List({
	@FieldMatch(first = "password", second = "matchingPassword", message = "The password fields must match")
})
public class UserDTO {
	
	@NotEmpty
	private String username;
	
	@ValidPassword
	private String password;
	
	private String matchingPassword;
	
	public UserDTO() {
		
	}
	
	public UserDTO(String username, String password, String matchingPassword) {
		this.username = username;
		this.password = password;
		this.matchingPassword = matchingPassword;
	}

	public String getUsername() {
		return username;
	}

	public void setUsername(String username) {
		this.username = username;
	}

	public String getPassword() {
		return password;
	}

	public void setPassword(String password) {
		this.password = password;
	}

	public String getMatchingPassword() {
		return matchingPassword;
	}

	public void setMatchingPassword(String matchingPassword) {
		this.matchingPassword = matchingPassword;
	}

	@Override
	public String toString() {
		return "UserDTO [username=" + username + ", password=" + password + ", matchingPassword=" + matchingPassword
				+ "]";
	}

}
