package com.excelportal.controller;

import java.util.Optional;

import javax.validation.Valid;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.propertyeditors.StringTrimmerEditor;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.WebDataBinder;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.InitBinder;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import com.excelportal.model.User;
import com.excelportal.model.UserDTO;
import com.excelportal.service.UserService;

@Controller
@RequestMapping("/register")
public class RegistrationController {
	
	@Autowired
	private UserService userService;
	
	@InitBinder
	public void initBinder(WebDataBinder dataBinder) {
		
		StringTrimmerEditor stringTrimmerEditor = new StringTrimmerEditor(true);
		
		dataBinder.registerCustomEditor(String.class, stringTrimmerEditor);
	}
	
	@GetMapping("/new-user")
	public String registerNewUser(Model model) {
		UserDTO userDTO = new UserDTO();
		model.addAttribute("userDTO", userDTO);
		return "registration-form";
	}
	
	@PostMapping("/process-new-user")
	public String registrationSuccessful(@Valid @ModelAttribute("userDTO") UserDTO userDTO, BindingResult bindingResult, Model model) {
		if(bindingResult.hasErrors()) {
			return "registration-form";
		}
		
		Optional<User> optionalUser = Optional.ofNullable(userService.findUserByUsername(userDTO.getUsername()));
		
		if(optionalUser.isPresent()) {
			model.addAttribute("newUser", new UserDTO());
			model.addAttribute("message", "Username already exists!");
			return "registration-form";
		}
		
		userService.saveUser(userDTO);
		
		return "registration-success";
	}
}
