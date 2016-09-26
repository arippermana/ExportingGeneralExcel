package my.mimos.tpcohcis.shared.util;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.json.simple.parser.ParseException;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.stream.Collector;
import java.util.stream.Collectors;

import javax.transaction.Transactional;

import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
@Transactional
public class TemplateForExcel {
	
	Logger LOG = LoggerFactory.getLogger(TemplateForExcel.class);
	
	public Map<String, ArrayList<Map<String, String>>> ExportExcel(MultipartFile file, String jsonTemplate, String headerTemplate){
		Map<String, String> excellTable = new HashMap<String, String>();
		ArrayList<Map<String, String>> allData = null;
		Map<String, ArrayList<Map<String, String>>> worksheet = new HashMap<String, ArrayList<Map<String, String>>>();
		Boolean isData = false;
		
		try{
			Workbook wb = new SXSSFWorkbook(200);
			wb = WorkbookFactory.create(file.getInputStream());
			//XSSFWorkbook wb = new XSSFWorkbook(file.getInputStream());
			DataFormatter objDefaultFormat = new DataFormatter();
			FormulaEvaluator objFormulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
			System.out.println("check formula evaluator : " + objFormulaEvaluator);
			ArrayList<String> nameTemplate = getHeaderExcel(jsonTemplate, headerTemplate);
			Sheet sheet = null;
			int maxCol = nameTemplate.size();
			
			for (int sheetAt = 0; sheetAt < wb.getNumberOfSheets(); sheetAt++){
				/*if (wb.getSheetName(sheetAt).contains("Sheet") || wb.getSheetName(sheetAt).contains("sheet")){
					continue;
				}*/
				
				allData = new ArrayList<Map<String, String>>();
				sheet = wb.getSheetAt(sheetAt);
				int rowNum = 0;
				for (Row row : sheet){
					excellTable = new HashMap<String, String>();
					for (int cn = 0; cn < row.getLastCellNum(); cn++){
						Cell cell = row.getCell(cn, Row.CREATE_NULL_AS_BLANK);
						// This will evaluate the cell, And any type of cell will return string value
						objFormulaEvaluator.evaluate(cell); 
						// end
					    String cellValueStr = objDefaultFormat.formatCellValue(cell,objFormulaEvaluator);
					    excellTable.put(nameTemplate.get(cn), cellValueStr);
						if (cn == row.getLastCellNum()-1){
							if (cn < (maxCol -1)){
								cn = cn + 1;
								for ( ;cn < maxCol; cn++){
									excellTable.put(nameTemplate.get(cn), "");
								}
							}
						}
					}
					
					Boolean isDataValid = false;
					
					
					for (Map.Entry<String, String> entry : excellTable.entrySet()){
						if (!StringUtils.isEmpty(entry.getValue())){
							isDataValid = true;
						}
					}
					
					if (isDataValid){
						excellTable.put("row", String.valueOf(rowNum));
						allData.add(excellTable);
					}
					rowNum += 1;
					//allData.add(excellTable);
					/*if (!MapUtils.isEmpty(excellTable)){
						allData.add(excellTable);
					}*/					
				}
				
				
				for (Map<String, String> checking : allData){
					for (Map.Entry<String, String> entry : checking.entrySet()){
						if (entry.getKey().contains(entry.getValue())){
							isData = true;
						}
					}
					
					if (isData){
						worksheet.put(wb.getSheetName(sheetAt), allData);
						continue;
					}
				}
				
				//System.out.println(worksheet.toString());
			}
			wb.close();
		} catch (Exception e){
			
		}
		System.out.println(worksheet.toString());
		return worksheet;
	}
	
	public ArrayList<String> getHeaderExcel(String jsonTemplate, String headerTemplate){
		ArrayList<String> headerFormat = new ArrayList<String>();
		JSONParser parser = new JSONParser();
		JSONObject jsonObj = null;
		ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
		
		try {
			//InputStream is = classLoader.getResourceAsStream(jsonTemplate);
			InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("templates" + File.separator + jsonTemplate);
			String line = "";
			try (BufferedReader buffer = new BufferedReader(new InputStreamReader(is))){
				line = buffer.lines().collect(Collectors.joining("\n"));
			}
			/*BufferedReader reader = new BufferedReader(new InputStreamReader(is));
	        StringBuilder out = new StringBuilder();
	        while ((line = reader.readLine()) != null) {
	            out.append(line);
	        }
	        //System.out.println(out.toString());   //Prints the string content read from input stream
	        reader.close();*/
	        
			jsonObj = (JSONObject) parser.parse(line);
		} catch (IOException | ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			System.out.println(e.toString());
		}
		JSONArray jsonFormat = (JSONArray) jsonObj.get(headerTemplate);

		Iterator<String> iter = jsonFormat.iterator();
		while(iter.hasNext()){
			headerFormat.add((String) iter.next());
		}
		return headerFormat;
	}
}
