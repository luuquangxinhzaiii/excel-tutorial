package com.example.demo.controller;

import com.example.demo.helper.excelHelper;
import com.example.demo.model.Cars;
import com.example.demo.service.CarService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@RestController
@CrossOrigin(origins = "*")
public class CarController {
    @Autowired
    CarService carService;

    @PostMapping("/car")
    public ResponseEntity<?> add(@RequestBody Cars car){
        car = carService.add(car);
        return new ResponseEntity(car, HttpStatus.OK);
    }

    @PutMapping("/car")
    public ResponseEntity<?> update(@RequestBody Cars car){
        car = carService.update(car);
        return new ResponseEntity(car, HttpStatus.OK);
    }

    @DeleteMapping("/car")
    public ResponseEntity<?> delete(@PathVariable(value = "id") long id){
        carService.delete(id);
        return new ResponseEntity("delete success", HttpStatus.OK);
    }

    @PostMapping("/upload")
    public ResponseEntity<?> uploadFile(@RequestParam("file") MultipartFile file) {
        String message = "";

        if (excelHelper.hasExcelFormat(file)) {
            try {
                carService.save(file);
                message = "Uploaded the file successfully: " + file.getOriginalFilename();
                return ResponseEntity.status(HttpStatus.OK).body(message);
            } catch (Exception e) {
                message = "Could not upload the file: " + file.getOriginalFilename() + "!";
                return ResponseEntity.status(HttpStatus.EXPECTATION_FAILED).body(message);
            }
        }

        message = "Please upload an excel file!";
        return ResponseEntity.status(HttpStatus.BAD_REQUEST).body(message);
    }

    @GetMapping("/tutorials")
    public ResponseEntity<?> getAllTutorials() {
        try {
            List<Cars> cars = carService.getAllTutorials();

            if (cars.isEmpty()) {
                return new ResponseEntity<>(HttpStatus.NO_CONTENT);
            }

            return new ResponseEntity<>(cars, HttpStatus.OK);
        } catch (Exception e) {
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @GetMapping("/download")
    public ResponseEntity<?> getFile() {
        String filename = "listCars.xlsx";
        InputStreamResource file = new InputStreamResource(carService.load());

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename)
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .body(file);
    }
}
