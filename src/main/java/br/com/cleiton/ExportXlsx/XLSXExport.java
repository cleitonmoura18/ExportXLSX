package br.com.cleiton.ExportXlsx;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


/**
 * Cria um arquivo XLSX passando uma coleção de elementos
 * @author Cleiton
 *
 * @param <T>
 */

public class XLSXExport<T> {
	private List<T> objects;
	private List<Element> elements;
	private Workbook workBook;
	/**
	 * Salva o arquivo a cada quantidade de linhas informa
	 */
	private Integer quantidadeLinhasParaSalvarArquivo =1000;
	
	/**
	 * Adiciona uma lista de elementos à lista que será exportada para csv
	 * @param List<T> - objects
	 * @return XLSXExport - XLSXExport
	 */
	public XLSXExport<T> addAll( List<T> objects ) {
		this.objects.addAll(objects);
		return this;
	}

	public XLSXExport() {
		this.objects = new ArrayList<T>();
		this.elements = new ArrayList<Element>();
	}

	public XLSXExport(List<T> objects) {
		this.objects = objects;
		this.elements = new ArrayList<Element>();
	}

	/**
	 * Adiciona um elemento à lista que será exportada para csv 
	 * @param Object - object
	 * @return XLSXExport - XLSXExport
	 */
	public XLSXExport<T> add( T object ) {
		this.objects.add(object);
		return this;
	}
	
	
	public XLSXExport<T> addColumn( String field, String displayName,  TipoColuna tipoColuna ) {
		Element e = new Element(field, displayName, tipoColuna);
		this.elements.add(e);
		return this;
	}
	
	/**
	 * Adiciona um par field(campo que deve ser apresentado no corpo do csv)
	 * dysplayName(Nome do header referente a este field) para representar uma
	 * Column no csv
	 * @param String - field é o campo que deve ser buscado via getter
	 * @param String - displayName é o nome no header referente a esta Coluna
	 * @return XLSXExport - XLSXExport
	 */
	public XLSXExport<T> addColumn( String field, String displayName ) {
		return addColumn(field, displayName, null);
	}
	
	/**
	 * Adiciona uma Coluna no csv
	 * @param String - field ( é o campo que deve ser buscado via getter)
	 * @return CSVExporter - CSVExporter
	 */
	public XLSXExport<T> addColumn( String field ) {
		Element e = new Element(field);
		this.elements.add(e);
		return this;
	}

	/**
	 * Gera o csv e retorna o arquivo gerado
	 * @param nomeAba 
	 * @return File
	 * @throws IOException
	 */
	public byte[] build(String nomeAba) throws IOException {
		if ( elements.isEmpty() ) {
			throw (new RuntimeException("Ao menos uma coluna deve ser configurada"));
		}

		workBook = new  SXSSFWorkbook(getQuantidadeLinhasParaSalvarArquivo());
		buildSheet(workBook, nomeAba);
		
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		workBook.write(byteArrayOutputStream);
		byteArrayOutputStream.close();
		return byteArrayOutputStream.toByteArray();
	}
	
	public static byte[] build(@SuppressWarnings("rawtypes") Map<String, XLSXExport> sheets) throws IOException {
		return build(sheets,1000);
	}
	/*
	 * Gerar um arquivo xlsx com mais de uma planilha.
	 * 
	 */
	@SuppressWarnings("rawtypes")
	public static byte[] build(Map<String, XLSXExport> sheets, int linhas) throws IOException {
		if ( sheets.isEmpty() ) {
			throw (new RuntimeException("Não foi informado nenhuma planilha"));
		}
		Workbook workBook = new SXSSFWorkbook(linhas);
		sheets.keySet().forEach( nome -> {			
			sheets.get(nome).buildSheet(workBook, nome);
		});
		
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		workBook.write(byteArrayOutputStream);
		byteArrayOutputStream.close();
		return byteArrayOutputStream.toByteArray();
	}
	
	private Sheet buildSheet(Workbook workBook, String nomeAba){
		Sheet sheet = workBook.createSheet(nomeAba);
		CellStyle cellStyleDate = workBook.createCellStyle();
		cellStyleDate.setDataFormat(workBook.createDataFormat().getFormat("dd/MM/yyyy"));

		CellStyle cellStyleDinheiro = workBook.createCellStyle();
		cellStyleDinheiro.setDataFormat((short) 8);
		
		Row currentRow = sheet.createRow(0);
		buildHeader(currentRow);
		for (int i = 0; i < objects.size(); i++) {
			Row createRow = sheet.createRow(i+1);
			writeLine(objects.get(i),createRow,cellStyleDate,cellStyleDinheiro);
			
		}
		return sheet;
	}

	/**
	 * Escreve o header do arquivo de acordo com os displayName especificados
	 * @param currentRow 
	 */
	private void buildHeader(Row currentRow) {
		for (int i = 0; i < elements.size(); i++) {
			Cell createCell = currentRow.createCell(i);
			createCell.setCellValue(elements.get(i).getDisplayName());
		}
		
	}

	/**
	 * Escreve uma linha no arquivo referente ao o objeto passado
	 * @param object
	 * @param createRow 
	 * @param cellStyleDinheiro 
	 * @param cellStyleDate 
	 */
	private void writeLine(T object, Row createRow, CellStyle cellStyleDate, CellStyle cellStyleDinheiro) {
		for (int i = 0; i < elements.size(); i++) {
			Element element = elements.get(i);
			Object value = UtilsReflection.getMethodCallResult(element, object);
			Cell createCell = createRow.createCell(i);
			TipoColuna tipoColuna = element.getTipoColuna();
			if(value == null){
				continue;
			}
			if(tipoColuna == null){
				createCell.setCellValue(value.toString());
				continue;
			}
			switch (tipoColuna) {
			case NUMERO:
				Double valor = new Double(value.toString().replace(".", "").replace(",", "."));
				createCell.setCellValue(valor);
				break;
			case DATA:
				Calendar calendar = createCalendarFromDate((Date) value);
				createCell.setCellValue(calendar);
				createCell.setCellStyle(cellStyleDate);
				break;
			case MONEY:
				createCell.setCellValue(new Double(value.toString()));
				createCell.setCellStyle(cellStyleDinheiro);
				break;
		
			default:
				createCell.setCellValue(value.toString());
				break;
			}

		}
	}
	
	
	public static Calendar createCalendarFromDate(Date date) {
		Calendar result = Calendar.getInstance();
		result.setTime(date);
		return result;
	}

	
	public Integer getQuantidadeLinhasParaSalvarArquivo() {
		return quantidadeLinhasParaSalvarArquivo;
	}

	public void setQuantidadeLinhasParaSalvarArquivo(Integer quantidadeLinhasParaSalvarArquivo) {
		this.quantidadeLinhasParaSalvarArquivo = quantidadeLinhasParaSalvarArquivo;
	}
	
}
