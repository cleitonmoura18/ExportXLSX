package br.com.cleiton.ExportXlsx;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.junit.Test;

public class XLSXExportTest {

	@Test
	public void test() throws IOException {
		List<Item> itens = criarItem();
		List<Item> itens2 = criarITem2();
		XLSXExport<Item> xlsxExport2 = criarArquivoXlsx2(itens2);
		HashMap<String, XLSXExport> mapas = new HashMap<>();
		mapas.put("Arquivo 1", criarArquivoXLSX(itens));
		mapas.put("Arquivo 2", xlsxExport2);
		byte[] build = XLSXExport.build(mapas,1000);
		/**
		 * caminho do local que vai ser salvo o arquivo
		 */
//		Files.write(Paths.get("/home/ihealth1/Documentos/Trabalho/eclipse/a.xlsx"), build);

	}

	private XLSXExport<Item> criarArquivoXlsx2(List<Item> itens2) {
		XLSXExport<Item> xlsxExport2 = new XLSXExport<Item>(itens2);
		xlsxExport2.addColumn("nome", "Nome").addColumn("valor", "Valor", TipoColuna.MONEY);
		return xlsxExport2;
	}

	private XLSXExport<Item> criarArquivoXLSX(List<Item> itens) {
		XLSXExport<Item> xlsxExport = new XLSXExport<Item>(itens);
		xlsxExport.setQuantidadeLinhasParaSalvarArquivo(1000);
		xlsxExport.addColumn("nome", "Nome").addColumn("valor", "Valor", TipoColuna.MONEY);
		return xlsxExport;
	}

	private List<Item> criarITem2() {
		List<Item> itens2 = new ArrayList<>();
		itens2.add(new Item("c", 10));
		itens2.add(new Item("d", 10));
		itens2.add(new Item("e", 10));
		return itens2;
	}

	private List<Item> criarItem() {
		List<Item> itens = new ArrayList<>();
		itens.add(new Item("a", 10));
		itens.add(new Item("b", 10));
		itens.add(new Item("v", 10));
		return itens;
	}

}
