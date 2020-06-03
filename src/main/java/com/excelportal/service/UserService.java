package com.excelportal.service;

import org.springframework.security.core.userdetails.UserDetailsService;

import com.excelportal.model.User;
import com.excelportal.model.UserDTO;

public interface UserService extends UserDetailsService {
	
	void saveUser(UserDTO userDTO);
	User findUserByUsername(String username);
}
