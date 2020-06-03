package com.excelportal.repository;

import java.util.Optional;

import com.excelportal.model.User;

public interface UserRepositoryCustom {
	
	User findUserByUsername(String username);
}
