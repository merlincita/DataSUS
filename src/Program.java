import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class Program {

	public static void main(String[] args) {
		int width = (int)(25 * 1.14388) * 256;
		//String input = "C:\\Users\\merlin\\workspace\\DataSUS\\lib\\UniqueID.xlsx";
		String input = "C:\\Users\\merlin\\workspace\\DataSUS\\lib\\2018UniqueIDs.xlsx";
		//GenerateModuloBasico(width, input, "C:/Users/merlin/workspace/DataSUS/lib/dataSus2017.xls");
		//System.out.println("The excel file 1 has been generated!");
		//GenerateCaracterizacao(width, input, "C:/Users/merlin/workspace/DataSUS/lib/dataSus3Col4rowsFails.xls" );
        //System.out.println("The excel file 2 has been generated!");
		
		GenerateData1File(width, input, "C:/Users/merlin/workspace/DataSUS/lib/dataSus2018_1.xls");
		System.out.println("The excel file has been generated!");
	}
	
	private static void GenerateData1File(int width, String input, String filename) {
		try {			
			FileInputStream inputFile = new FileInputStream(new File(input));
            
			//Get the workbook instance for XLS file 
			XSSFWorkbook workbookInput = new XSSFWorkbook(inputFile);
			
			//Get sheet from the workbook
			XSSFSheet sheetInput = workbookInput.getSheetAt(0);
			 
			//Get iterator to all the rows in current sheet
			Iterator<Row> rowIterator = sheetInput.iterator();
			
			short index = 0;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Informacao");
			HSSFRow rowhead = sheet.createRow(index);
            rowhead.createCell(0).setCellValue("CNES");
            rowhead.createCell(1).setCellValue("Estabelecimento");
            rowhead.createCell(2).setCellValue("Nome Empresarial");
            rowhead.createCell(3).setCellValue("Logradouro");
            rowhead.createCell(4).setCellValue("Numero");
            rowhead.createCell(5).setCellValue("Complemento");
            rowhead.createCell(6).setCellValue("Bairro");
            rowhead.createCell(7).setCellValue("CEP");
            rowhead.createCell(8).setCellValue("Municipio");
            rowhead.createCell(9).setCellValue("UF");
            rowhead.createCell(10).setCellValue("Fax");
            rowhead.createCell(11).setCellValue("Telefone");
            rowhead.createCell(12).setCellValue("Email");
            rowhead.createCell(13).setCellValue("CNPJ");
            rowhead.createCell(14).setCellValue("CPF");
            rowhead.createCell(15).setCellValue("CNPJ Mantenedora");
            rowhead.createCell(16).setCellValue("Diretor Clinico");
            rowhead.createCell(17).setCellValue("Representante Legal");
            rowhead.createCell(18).setCellValue("Representante Legal Cargo");
            rowhead.createCell(19).setCellValue("Representante Legal Email");
            
            rowhead.createCell(20).setCellValue("Tipo Estabelecimento");
            rowhead.createCell(21).setCellValue("Atividade Ensino/Pesquisa");
            rowhead.createCell(22).setCellValue("Codigo/Natureza Jurídica");
            
            sheet.setColumnWidth(1, width);
            sheet.setColumnWidth(2, width);
            sheet.setColumnWidth(3, width);
            sheet.setColumnWidth(16, width);
            sheet.setColumnWidth(17, width);
            
            sheet.setColumnWidth(20, width);
            sheet.setColumnWidth(21, width);
            sheet.setColumnWidth(22, width * 2);
            
            rowIterator.next(); // to skip the title header
            Document doc;
            Element table;
            HSSFRow row = null;
            Elements rows;
            boolean fine;
			while (rowIterator.hasNext()) {
				//if (index == 5) break;
				index ++;
				fine = true;
				Row inputRow = rowIterator.next();
				int cnesInput = (int)inputRow.getCell(0).getNumericCellValue();
				//int cnesInput = Integer.parseInt(inputRow.getCell(0).getStringCellValue());	
				String cnes = String.format("%07d", cnesInput);
				String ufColumn = "" + (int)inputRow.getCell(1).getNumericCellValue();
				//String ufColumn = inputRow.getCell(1).getStringCellValue();
				try {
					// need http protocol
					doc = Jsoup.connect("http://cnes2.datasus.gov.br/Mod_Basico.asp?VCo_Unidade=" + ufColumn + cnes).get();
		
					Element estabelimento = doc.select("font[color='#ffcc99'][face='Verdana,arial'][size='1']").first();
					
					// get the specific table
					table = doc.select("table[bgcolor='white']").first();
					//Elements rows = table.select("td:not(bgcolor) > font[color='#003366'][face='Verdana,arial'][size='1']");
					rows = table.select("td:not(bgcolor) > font");
					String nomeempresarial = rows.get(8).text();
					String logradouro = rows.get(9).text();
					String numero = rows.get(16).text();
					String complemento = rows.get(17).text();
					String bairro = rows.get(18).text();
					String cep = rows.get(19).text();
					String municipio = rows.get(20).text();
					String uf = rows.get(21).text();
					String fax = rows.get(31).text();
					String telefone = rows.get(37).text();
					String email = rows.get(38).text();
					String cnpj = rows.get(39).text();
					String cpf = rows.get(40).text();
					String cnpjmanejadora = rows.get(41).text();
					String diretorclinico = rows.get(43).text();
					String reprlegalnome = rows.get(48).text(); // representante legal
					String replegalcargo = rows.get(49).text();
					String replegalemail = rows.get(50).text();					        
		            
		            row = sheet.createRow(index);
		            row.createCell(0).setCellValue(cnesInput);
		            row.createCell(1).setCellValue(estabelimento.text());
		            row.createCell(2).setCellValue(nomeempresarial);
		            row.createCell(3).setCellValue(logradouro);
		            row.createCell(4).setCellValue(numero);
		            row.createCell(5).setCellValue(complemento);
		            row.createCell(6).setCellValue(bairro);
		            row.createCell(7).setCellValue(cep);
		            row.createCell(8).setCellValue(municipio);
		            row.createCell(9).setCellValue(uf);
		            row.createCell(10).setCellValue(fax);
		            row.createCell(11).setCellValue(telefone);
		            row.createCell(12).setCellValue(email);
		            row.createCell(13).setCellValue(cnpj);
		            row.createCell(14).setCellValue(cpf);
		            row.createCell(15).setCellValue(cnpjmanejadora);
		            row.createCell(16).setCellValue(diretorclinico);
		            row.createCell(17).setCellValue(reprlegalnome);
		            row.createCell(18).setCellValue(replegalcargo);
		            row.createCell(19).setCellValue(replegalemail);
				}				
				catch(Exception exxx){
					System.out.println("Fail cnes " + cnes + " * " + ufColumn);
					index --;
					fine = false;
				}
				try
				{
		            doc = Jsoup.connect("http://cnes2.datasus.gov.br/Mod_Bas_Caracterizacao.asp?VCo_Unidade=" + ufColumn + cnes).get();
		    		
					// get the specific table
					table = doc.select("table[bgcolor='white'][cellpadding='1']").first();					
					if (!fine)
						row = sheet.createRow(index);
					
					//Elements rows = table.select("td > font[size='1'][face='Verdana,arial'][color='#003366']");
					rows = table.select("td > font");
					String tipoEstab = rows.get(2).text();
					String atividadeEnsino = rows.get(3).text();
					String codnatureza = rows.get(5).text();
		            
		            row.createCell(20).setCellValue(tipoEstab);
		            row.createCell(21).setCellValue(atividadeEnsino);
		            row.createCell(22).setCellValue(codnatureza);
				}
			
				catch(Exception exxx){
					System.out.println("Fail mbas " + cnes + " * " + ufColumn);
					if (fine)
						index --;
				}
			}
			FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            //workbookInput.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	
	

	private static void GenerateModuloBasico(int width, String input, String filename) {
		try {			
			FileInputStream inputFile = new FileInputStream(new File(input));
            
			//Get the workbook instance for XLS file 
			XSSFWorkbook workbookInput = new XSSFWorkbook(inputFile);
			
			//Get sheet from the workbook
			XSSFSheet sheetInput = workbookInput.getSheetAt(0);
			 
			//Get iterator to all the rows in current sheet
			Iterator<Row> rowIterator = sheetInput.iterator();
			
			short index = 0;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Informacao");
			HSSFRow rowhead = sheet.createRow(index);
            rowhead.createCell(0).setCellValue("CNES");
            rowhead.createCell(1).setCellValue("Estabelecimento");
            rowhead.createCell(2).setCellValue("Nome Empresarial");
            rowhead.createCell(3).setCellValue("Logradouro");
            rowhead.createCell(4).setCellValue("Numero");
            rowhead.createCell(5).setCellValue("Complemento");
            rowhead.createCell(6).setCellValue("Bairro");
            rowhead.createCell(7).setCellValue("CEP");
            rowhead.createCell(8).setCellValue("Municipio");
            rowhead.createCell(9).setCellValue("UF");
            rowhead.createCell(10).setCellValue("Fax");
            rowhead.createCell(11).setCellValue("Telefone");
            rowhead.createCell(12).setCellValue("Email");
            rowhead.createCell(13).setCellValue("CNPJ");
            rowhead.createCell(14).setCellValue("CPF");
            rowhead.createCell(15).setCellValue("CNPJ Mantenedora");
            rowhead.createCell(16).setCellValue("Diretor Clinico");
            rowhead.createCell(17).setCellValue("Representante Legal");
            rowhead.createCell(18).setCellValue("Representante Legal Cargo");
            rowhead.createCell(19).setCellValue("Representante Legal Email");
            
            sheet.setColumnWidth(1, width);
            sheet.setColumnWidth(2, width);
            sheet.setColumnWidth(3, width);
            sheet.setColumnWidth(16, width);
            sheet.setColumnWidth(17, width);
            
            //rowIterator.next(); // to skip the title header
			while (rowIterator.hasNext()) {
				index ++;
				Row inputRow = rowIterator.next();
				int cnesInput = (int)inputRow.getCell(0).getNumericCellValue();	
				String cnes = String.format("%07d", cnesInput);
				String ufColumn = "" + (int)inputRow.getCell(1).getNumericCellValue();
				
				try {
					// need http protocol
					Document doc = Jsoup.connect("http://cnes2.datasus.gov.br/Mod_Basico.asp?VCo_Unidade=" + ufColumn + cnes).get();
		
					Element estabelimento = doc.select("font[color='#ffcc99'][face='Verdana,arial'][size='1']").first();
					
					// get the specific table
					Element table = doc.select("table[bgcolor='white']").first();
					Elements rows = table.select("td:not(bgcolor) > font[color='#003366'][face='Verdana,arial'][size='1']");
					String nomeempresarial = rows.get(8).text();
					String logradouro = rows.get(9).text();
					String numero = rows.get(16).text();
					String complemento = rows.get(17).text();
					String bairro = rows.get(18).text();
					String cep = rows.get(19).text();
					String municipio = rows.get(20).text();
					String uf = rows.get(21).text();
					String fax = rows.get(31).text();
					String telefone = rows.get(37).text();
					String email = rows.get(38).text();
					String cnpj = rows.get(39).text();
					String cpf = rows.get(40).text();
					String cnpjmanejadora = rows.get(41).text();
					String diretorclinico = rows.get(43).text();
					String reprlegalnome = rows.get(48).text(); // representante legal
					String replegalcargo = rows.get(49).text();
					String replegalemail = rows.get(50).text();
					/*for (byte i = 0; i < rows.size(); i ++) {
						Element link = rows.get(i);
						System.out.println(i + " *** " + link.text());
					}*/	            
		            
		            HSSFRow row = sheet.createRow(index);
		            row.createCell(0).setCellValue(cnesInput);
		            row.createCell(1).setCellValue(estabelimento.text());
		            row.createCell(2).setCellValue(nomeempresarial);
		            row.createCell(3).setCellValue(logradouro);
		            row.createCell(4).setCellValue(numero);
		            row.createCell(5).setCellValue(complemento);
		            row.createCell(6).setCellValue(bairro);
		            row.createCell(7).setCellValue(cep);
		            row.createCell(8).setCellValue(municipio);
		            row.createCell(9).setCellValue(uf);
		            row.createCell(10).setCellValue(fax);
		            row.createCell(11).setCellValue(telefone);
		            row.createCell(12).setCellValue(email);
		            row.createCell(13).setCellValue(cnpj);
		            row.createCell(14).setCellValue(cpf);
		            row.createCell(15).setCellValue(cnpjmanejadora);
		            row.createCell(16).setCellValue(diretorclinico);
		            row.createCell(17).setCellValue(reprlegalnome);
		            row.createCell(18).setCellValue(replegalcargo);
		            row.createCell(19).setCellValue(replegalemail);
				}
			
				catch(Exception exxx){
					System.out.println("Fail cnes " + cnes + " * " + ufColumn);
					index --;
				}
			}
			FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            workbookInput.close();
		} catch (IOException e) {
			e.printStackTrace();
			//return;
		}
	}

	private static void GenerateCaracterizacao(int width, String input, String filename) {
		try {
			FileInputStream inputFile = new FileInputStream(new File(input));
            
			//Get the workbook instance for XLS file 
			XSSFWorkbook workbookInput = new XSSFWorkbook(inputFile);
			
			//Get sheet from the workbook
			XSSFSheet sheetInput = workbookInput.getSheetAt(0);
			 
			//Get iterator to all the rows in current sheet
			Iterator<Row> rowIterator = sheetInput.iterator();
			
			short index = 0;
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Informacao");
			HSSFRow rowhead = sheet.createRow(index);
            rowhead.createCell(0).setCellValue("CNES");
            rowhead.createCell(1).setCellValue("Tipo Estabelecimento");
            rowhead.createCell(2).setCellValue("Atividade Ensino/Pesquisa");
            rowhead.createCell(3).setCellValue("Codigo/Natureza Jurídica");
            
            sheet.setColumnWidth(1, width);
            sheet.setColumnWidth(2, width);
            sheet.setColumnWidth(3, width * 2);
            
            //rowIterator.next(); // to skip the title header
			while (rowIterator.hasNext()) {
				index ++;
				Row inputRow = rowIterator.next();
				int cnesInput = (int)inputRow.getCell(0).getNumericCellValue();	
				String cnes = String.format("%07d", cnesInput);
				String ufColumn = "" + (int)inputRow.getCell(1).getNumericCellValue();
				
				try {
					// need http protocol
					Document doc = Jsoup.connect("http://cnes2.datasus.gov.br/Mod_Bas_Caracterizacao.asp?VCo_Unidade=" + ufColumn + cnes).get();
		
					// get the specific table
					Element table = doc.select("table[bgcolor='white'][cellpadding='1']").first();
					Elements rows = table.select("td > font[size='1'][face='Verdana,arial'][color='#003366']");
					String tipoEstab = rows.get(2).text();
					String atividadeEnsino = rows.get(3).text();
					String codnatureza = rows.get(5).text();
		            
		            HSSFRow row = sheet.createRow(index);
		            row.createCell(0).setCellValue(cnesInput);
		            row.createCell(1).setCellValue(tipoEstab);
		            row.createCell(2).setCellValue(atividadeEnsino);
		            row.createCell(3).setCellValue(codnatureza);
				}
			
				catch(Exception exxx){
					System.out.println("Fail 3 cnes " + cnes + " * " + ufColumn);
					//index --;
				}
			}
			FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            workbookInput.close();
		} catch (IOException e) {
			e.printStackTrace();
			//return;
		}
	}

}