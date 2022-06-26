package com.example.demo.service;

import com.example.demo.helper.excelHelper;
import com.example.demo.model.Cars;
import com.example.demo.repository.carRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

@Service
public class CarService {
    @Autowired
    carRepository carRepository;
    public Cars add(Cars cars){
        cars =  carRepository.save(cars);
        return cars;
    }

    public Cars update(Cars cars){
        cars = carRepository.save(cars);
        return cars;
    }

    public void delete(long id){
        carRepository.deleteById(id);
    }

    public void save(MultipartFile file) {
        try {
            List<Cars> tutorials = excelHelper.excelToCars(file.getInputStream());
            carRepository.saveAll(tutorials);
        } catch (IOException e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
    }

    public ByteArrayInputStream load() {
        List<Cars> tutorials = carRepository.findAll();

        ByteArrayInputStream in = excelHelper.carsToExcel(tutorials);
        return in;
    }

    public List<Cars> getAllTutorials() {
        return carRepository.findAll();
    }
}
