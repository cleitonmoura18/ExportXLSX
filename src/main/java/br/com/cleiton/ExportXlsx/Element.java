package br.com.cleiton.ExportXlsx;

/**
 * @author Djair
 * Classe pra encapsular o nome a ser utilizado no header do csv e o e o nome do getter referente ao campo
 * @version 0.0.2
 */
public class Element {
	
	private String field;
	private String displayName;
	private TipoColuna tipoColuna;
	
	public Element(String field, String displayName) {
		if(field == null || field.isEmpty()){
			throw new RuntimeException("field can't be blank");
		}
		setField(field);
		setDisplayName(displayName);
	}
	
	/**
	 * @param field
	 * @param displayName
	 * @param nullable
	 */
	public Element(String field, String displayName, TipoColuna tipoColuna) {
		this(field, displayName);
		this.tipoColuna = tipoColuna;
	}

	public Element(String field) {
		if(field == null || field.isEmpty()){
			throw new RuntimeException("field can't be blank");
		}
		setField(field);
		setDefaultDisplayName(field);
	}
	
	private void setDefaultDisplayName( String fieldName ) {
		String displayName = fieldName;
		if(fieldName.contains(".")){
			int beginOfLastField = displayName.lastIndexOf(".") + 1;
			displayName = displayName.substring(beginOfLastField);
		}
		this.displayName = displayName.substring(0, 1).toUpperCase() + displayName.substring(1);
	}

	public String getField() {
		return field;
	}
	public void setField( String field ) {
		this.field = field;
	}
	public String getDisplayName() {
		return displayName;
	}
	public void setDisplayName( String displayName ) {
		this.displayName = displayName;
	}	
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((displayName == null) ? 0 : displayName.hashCode());
		result = prime * result + ((field == null) ? 0 : field.hashCode());
		return result;
	}
	@Override
	public boolean equals( Object obj ) {
		if ( this == obj )
			return true;
		if ( obj == null )
			return false;
		if ( getClass() != obj.getClass() )
			return false;
		Element other = (Element) obj;
		if ( displayName == null ) {
			if ( other.displayName != null )
				return false;
		} else if ( !displayName.equals(other.displayName) )
			return false;
		if ( field == null ) {
			if ( other.field != null )
				return false;
		} else if ( !field.equals(other.field) )
			return false;
		return true;
	}

	public TipoColuna getTipoColuna() {
		return tipoColuna;
	}

	public void setTipoColuna(TipoColuna tipoColuna) {
		this.tipoColuna = tipoColuna;
	}
}
