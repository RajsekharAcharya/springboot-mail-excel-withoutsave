package com.example.maildemo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.scheduling.annotation.EnableScheduling;

@SpringBootApplication
@EnableScheduling
public class MaildemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(MaildemoApplication.class, args);
	}

}
