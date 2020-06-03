package com.excelportal.repository;

import org.hibernate.Session;
import org.hibernate.query.Query;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.transaction.annotation.Transactional;

import javax.persistence.EntityManager;


import com.excelportal.model.Role;

public class RoleRepositoryCustomImpl implements RoleRepositoryCustom {
	
	@Autowired
	private EntityManager entityManager;
	
	@Override
	@Transactional
	public Role findRoleByType(String roleType) {
		Session session = entityManager.unwrap(Session.class);
		Query<Role> query = session.createQuery("from Role where type=:roleType", Role.class);
		Role retrievedRole;
		try {
			retrievedRole = query.getSingleResult();
		} catch(Exception e) {
			e.printStackTrace();
			retrievedRole = null;
		}
		return retrievedRole;
	}

}
