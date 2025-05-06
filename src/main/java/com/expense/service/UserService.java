package com.expense.service;



import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.stereotype.Service;

import com.expense.model.User;
import com.expense.repository.UserRepository;

@Service
public class UserService {
	
	@Autowired
	private UserRepository userRepository;
	
	@Autowired
	private BCryptPasswordEncoder bCryptPasswordEncoder;
	
	public void add(User user)
	{	
		user.setPassword(bCryptPasswordEncoder.encode(user.getPassword()));
		userRepository.save(user);
	}
	
	public List<User> getAll()
	{
		return userRepository.findAll();
	}

}

