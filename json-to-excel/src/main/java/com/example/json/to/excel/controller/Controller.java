package com.example.json.to.excel.controller;

import com.mongodb.client.MongoClient;
import com.mongodb.client.MongoClients;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;
import jdk.security.jarsigner.JarSigner;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@RestController
public class Controller {

    @Autowired
    private ResourceLoader resourceLoader;

    MongoClient mongoClient= MongoClients.create();
    MongoDatabase mongoDatabase= mongoClient.getDatabase("my-movies");

    MongoCollection<Document> mongoCollection= mongoDatabase.getCollection("movies-list");

    @GetMapping("/csvToMongoDB")
    public void csvToMongoDB() throws IOException, CsvValidationException {
        LocalTime localTime=LocalTime.now();
        System.out.println(localTime);
        CSVReader csvReader=new CSVReader(new FileReader("movies.csv"));
        String[] header= csvReader.readNext();
        int j=1;
        List<Document> documents = new ArrayList<>();

        String[] values;
        while((values=csvReader.readNext())!=null){
            Map<String,Object> data=new HashMap<>();
            for(int i=0;i<header.length;i++)
            {
                data.put(header[i],values[i]);
            }
            documents.add(new Document(data));
            j++;
        }
        mongoCollection.insertMany(documents);
        System.out.println("Total records saved are: "+j);
        System.out.println(LocalTime.now());

    }

    @GetMapping("/convertMongoDBToExcel")
    // iterating by storing all the data in list such that iteration will be easy
    public void convertMongoDBToExcel() throws IOException {

        System.out.println("Headers Insertion \n"+LocalTime.now());

        List<Document> documentList=mongoCollection.find().into(new ArrayList<>());

        Document header=mongoCollection.find().first();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data");

        int rowCount=0;
        Row row=sheet.createRow(rowCount++);
        int colCount=0;
        for(String head: header.keySet()){
            if (!(head.equals("_id"))){
                continue;
            }
            Cell cell=row.createCell(colCount++);
            cell.setCellValue(head);
        }
        System.out.println("Headers Inserted \n"+LocalTime.now());

        System.out.println("Values Insertion \n"+LocalTime.now());
        for(Document document:documentList){
            row=sheet.createRow(rowCount++);
            colCount=0;
            for(String key: document.keySet()){
                if (key.equals("_id"))
                {
                    continue;
                }
                Cell cell= row.createCell(colCount++);
                cell.setCellValue(document.getString(key));
            }
        }

        FileOutputStream fileOutputStream=new FileOutputStream("movies.xlsx");
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();

        System.out.println("Values Inserted \n"+LocalTime.now());

    }

    @GetMapping("/nestedJson")
    public void nestedJson() throws IOException {
        Resource resource = resourceLoader.getResource("classpath:students.json");
        InputStreamReader inputStreamReader = new InputStreamReader(resource.getInputStream());
        String jsonData = FileCopyUtils.copyToString(inputStreamReader);


        JSONObject jsonObject = new JSONObject(jsonData);
        System.out.println(jsonObject.keySet());
        System.out.println(jsonObject.getString("company"));
        String dbname=jsonObject.getString("company").replace(' ','_');
        jsonObject.remove("company");
        MongoDatabase mongoDatabase1=mongoClient.getDatabase(dbname);

        for(String s: jsonObject.keySet()){
            mongoDatabase1.createCollection(s);
            JSONArray jsonArray=jsonObject.optJSONArray(s);
            List<Document> data=new ArrayList<>();
            for(int i=0;i<jsonArray.length();i++){
                JSONObject values=jsonArray.getJSONObject(i);
                data.add(new Document(values.toMap()));
            }
            MongoCollection<Document> documentMongoCollection = mongoDatabase1.getCollection(s);
            documentMongoCollection.insertMany(data);
            System.out.println("Inserted Values of: "+s);
        }

    }

    @GetMapping("/jsonToExcelViaMongo")
    public void jsonToExcelViaMongo() throws IOException {
        System.out.println("Insertion Started \n"+LocalTime.now());
        Resource resource = resourceLoader.getResource("classpath:jsonFile.json");
        InputStreamReader inputStreamReader = new InputStreamReader(resource.getInputStream());
        String jsonData = FileCopyUtils.copyToString(inputStreamReader);

        List<Document> documents=new ArrayList<>();

        JSONObject jsonObject = new JSONObject(jsonData);

        for (String string:jsonObject.keySet()){
            documents.add(new Document(jsonObject.getJSONObject(string).toMap()));
        }
        mongoCollection.insertMany(documents);
        System.out.println("Insertion Done \n"+LocalTime.now());

        //*************************** to store in excel **********************************

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data");
        Document header=documents.get(0);
        int rowCount=0;
        Row row=sheet.createRow(rowCount++);
        int colCount=0;
        for(String head: header.keySet()){
            if (!(head.equals("_id"))){
                continue;
            }
            Cell cell=row.createCell(colCount++);
            cell.setCellValue(head);
        }
        System.out.println("Headers Inserted \n"+LocalTime.now());

        System.out.println("Values Insertion \n"+LocalTime.now());
        for(Document document:documents){
            row=sheet.createRow(rowCount++);
            colCount=0;
            for(String key: document.keySet()){
                if (key.equals("_id"))
                {
                    continue;
                }
                Cell cell= row.createCell(colCount++);
                cell.setCellValue(document.getString(key));
            }
        }

        FileOutputStream fileOutputStream=new FileOutputStream("movies.xlsx");
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();

        System.out.println("Values Inserted \n"+LocalTime.now());

    }

    @GetMapping("/mongoToExcel")
    public void mongoToExcel() throws IOException {
        Document header=mongoCollection.find().first();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data");

        int rowCount=0;
        Row row=sheet.createRow(rowCount++);
        int colCount=0;
        for(String head: header.keySet()){
            if (head.equals("_id")){
                continue;
            }
            Cell cell=row.createCell(colCount++);
            cell.setCellValue(head);
        }
        System.out.println("Headers Inserted");

        for(Document document:mongoCollection.find()){
            Row rows=sheet.createRow(rowCount++);
            colCount=0;
            document.remove("_id");

            for(String key: document.keySet()){

                Cell cell= rows.createCell(colCount++);
                cell.setCellValue(document.getString(key));
            }
        }

        FileOutputStream fileOutputStream=new FileOutputStream("movies.xlsx");
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();

        System.out.println(LocalTime.now());
        System.out.println("Values Inserted");
    }

    @GetMapping("/jsonToExcel")
    public void jsonToExcel() throws IOException {
        System.out.println(LocalTime.now());

        Resource resource = resourceLoader.getResource("classpath:jsonFile.json");
        InputStreamReader inputStreamReader = new InputStreamReader(resource.getInputStream());
        String jsonData = FileCopyUtils.copyToString(inputStreamReader);


        JSONObject jsonObject = new JSONObject(jsonData);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("data");

        int rowCount = 0;
        Row row = sheet.createRow(rowCount++);

        for (String s : jsonObject.keySet()) {
            int colCount = 0;
            JSONObject object = jsonObject.getJSONObject(s);
            for (String header : object.keySet()) {
                Cell cell = row.createCell(colCount++);
                cell.setCellValue(header);
            }
            break;
        }

        for (String s : jsonObject.keySet()) {
            Row row1 = sheet.createRow(rowCount++);
            int colCount = 0;
            JSONObject object = jsonObject.getJSONObject(s);
            for (String header : object.keySet()) {
                Cell cell = row1.createCell(colCount++);
                cell.setCellValue(object.getString(header));
            }
        }

        FileOutputStream outputStream = new FileOutputStream("movies.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

        System.out.println(LocalTime.now());
    }
}