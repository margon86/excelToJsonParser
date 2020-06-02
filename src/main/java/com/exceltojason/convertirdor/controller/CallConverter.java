package com.exceltojason.convertirdor.controller;

import java.util.ResourceBundle;

import org.apache.log4j.Logger;

import com.exceltojason.convertirdor.util.ParserExcelToJson;

public class CallConverter {
	
	final static Logger logger = Logger.getLogger(CallConverter.class);
	
	public static void main(String[] args) {
		
		logger.info("Iniciamos el programa excelToJason");
		
		//obtemos los paths desde el conf.properties
		ResourceBundle mybundle = ResourceBundle.getBundle("conf");
		String excelFilePath = mybundle.getString("excelPath");	
		String jsonFilePath = mybundle.getString("jsonPath");
		
		CallConverter call = new CallConverter();
		call.convertExcelToJson(excelFilePath ,jsonFilePath );
		

	}
	
	
	public void convertExcelToJson(String excelFilePath, String jsonFilePath ) {
		try {
			
			logger.info("Llamamos al parseador");
			ParserExcelToJson p = new ParserExcelToJson();
			String retorno = p.parseador(excelFilePath,jsonFilePath);
			if(retorno.equals("OK")) {
				logger.info("Proceso finalizado con exito");
			}else {
				logger.error(retorno);
			}
			
			
		} catch (Exception e) {
			logger.error("Error: "+e.getMessage());
		}
	}

}
