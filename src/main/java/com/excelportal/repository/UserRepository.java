package com.excelportal.repository;

import org.springframework.data.jpa.repository.JpaRepository;

import com.excelportal.model.User;

public interface UserRepository extends JpaRepository<User, Integer>, UserRepositoryCustom {
	
}
