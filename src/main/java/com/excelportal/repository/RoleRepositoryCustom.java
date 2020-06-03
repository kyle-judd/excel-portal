package com.excelportal.repository;

import com.excelportal.model.Role;

public interface RoleRepositoryCustom {
	
	Role findRoleByType(String roleType);
}
