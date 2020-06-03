package com.excelportal.repository;

import javax.persistence.EntityManager;

import org.hibernate.Session;
import org.hibernate.query.Query;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.transaction.annotation.Transactional;

import com.excelportal.model.User;

public class UserRepositoryCustomImpl implements UserRepositoryCustom {
	
	@Autowired
	private EntityManager entityManager;
	
	@Override
	@Transactional
	public User findUserByUsername(String username) {
		Session session = entityManager.unwrap(Session.class);
		Query<User> query = session.createQuery("from User where username=:username", User.class);
		query.setParameter("username", username);
		User retrievedUser;
		try {
			retrievedUser = query.getSingleResult();
		} catch(Exception e) {
			retrievedUser = null;
		}
		return retrievedUser;
	}

}
