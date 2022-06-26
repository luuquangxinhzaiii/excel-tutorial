package com.example.demo.repository;

import com.example.demo.model.Cars;
import org.springframework.data.jpa.repository.JpaRepository;

public interface carRepository extends JpaRepository<Cars, Long> {

}
