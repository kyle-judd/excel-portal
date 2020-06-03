package com.excelportal.service;

import java.util.Arrays;
import java.util.Collection;
import java.util.Optional;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.excelportal.model.Role;
import com.excelportal.model.User;
import com.excelportal.model.UserDTO;
import com.excelportal.repository.RoleRepository;
import com.excelportal.repository.UserRepository;

@Service
public class UserServiceImpl implements UserService {
	
	@Autowired
	private UserRepository userRepository;
	
	@Autowired
	private RoleRepository roleRepository;
	
	@Autowired
	private BCryptPasswordEncoder passwordEncoder;
	
	@Override
	public UserDetails loadUserByUsername(String username) throws UsernameNotFoundException {
		
		User retrievedUser = null;
		
		try {
			retrievedUser = userRepository.findUserByUsername(username);
		} catch(UsernameNotFoundException e) {
			e.printStackTrace();
		}
		return new org.springframework.security.core.userdetails.User(retrievedUser.getUsername(), retrievedUser.getPassword(), mapRolesToAuthorities(retrievedUser.getRoles()));
		
	}
	
	private Collection<? extends GrantedAuthority> mapRolesToAuthorities(Collection<Role> roles) {
		return roles.stream().map(role -> new SimpleGrantedAuthority(role.getType())).collect(Collectors.toList());
	}

	@Override
	@Transactional
	public void saveUser(UserDTO userDTO) {
		User newUser = new User();
		newUser.setUsername(userDTO.getUsername());
		newUser.setPassword(passwordEncoder.encode(userDTO.getPassword()));
		Optional<Role> optionalRole = roleRepository.findById(1);
		newUser.setRoles(Arrays.asList(optionalRole.get()));
		userRepository.save(newUser);
	}

	@Override
	public User findUserByUsername(String username) {
		return userRepository.findUserByUsername(username);
	}

}
