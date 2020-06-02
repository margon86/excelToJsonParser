package com.exceltojason.convertirdor.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;

public class ParserExcelToJson {
	
	final static Logger logger = Logger.getLogger(ParserExcelToJson.class);
	
	ParserCaracteresEspeciales parserCaracteres = new ParserCaracteresEspeciales();
	
	public String parseador(String pathExcel, String pathJson) {
		
		logger.info("Parseador pathExcel-->"+pathExcel);
		logger.info("Parseador pathJson-->"+pathJson);
		try {
			FileInputStream imp = new FileInputStream( new File(pathExcel) );
			Workbook workbook = WorkbookFactory.create( imp );

			//Obtenemos la primera hoja de la planilla.
			Sheet sheet = workbook.getSheetAt( 0 );

			    //Empezamos a construir el JSON.
			    JSONObject json = new JSONObject();

			    //Iteramos a traves de las filas.
			    JSONArray rows = new JSONArray();
			    for ( Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext(); )
			    {
			        Row row = rowsIT.next();
			        JSONObject jRow = new JSONObject();
			        //Iteramos a traves de las celdas.
			        JSONArray cells = new JSONArray();
			        for ( Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext(); )
			        {
			            Cell cell = cellsIT.next();
			            //Verificamos los tipos de datos de las celdas
			            if(cell.getCellType().toString().equals("STRING")) {
			            	cells.put( parserCaracteres.parsearCaracteresEspeciales(cell.getStringCellValue()) );
			            }else {
			            	//Al no ser string el tipo de datos es numeric
			            	//Los date tambien son de tipo numeric por lo tanto primero verificamos si es o no una fecha
			            	if (DateUtil.isCellDateFormatted(cell)){
			            	  SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
							  String cellValue = sdf.format(cell.getDateCellValue());
							  cells.put( cellValue );
			            	}else {
			            		Integer n = (int) cell.getNumericCellValue();
				            	cells.put( n.toString() );
			            	}
			            }
			            
			        }
			        //jRow.put( "cell", cells ); //--> en caso de que cada celda querramos almacenarlos como objeto json
			        rows.put( cells ); //--> guardamos todoa la fila como un objeto json
			    }

			    // Creamos el JSON.
			    json.put( "rows", rows );
			
			//Obtenemos el json generado y creamos archivo .json
			FileWriter file = new FileWriter(pathJson);
			file.write(json.toString(4)); // el 4 especifica que el string es en formato json
			file.close();
			logger.info("Proceso finalizado con exito!");
			return "OK";
		} catch (Exception e) {
			logger.error("Error: "+e.getMessage());
			return "Error: "+e.getMessage();
		}
	}

}
