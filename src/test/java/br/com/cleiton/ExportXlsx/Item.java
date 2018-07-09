package br.com.cleiton.ExportXlsx;

public class Item {
	private String nome;
	private Integer valor;

	public Item(String nome, Integer valor) {
		this.nome = nome;
		this.valor = valor;
	}

	public String getNome() {
		return nome;
	}

	public void setNome(String nome) {
		this.nome = nome;
	}

	public Integer getValor() {
		return valor;
	}

	public void setValor(Integer valor) {
		this.valor = valor;
	}

}
