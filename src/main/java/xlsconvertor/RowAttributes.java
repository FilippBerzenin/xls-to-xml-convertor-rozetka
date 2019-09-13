package xlsconvertor;

import lombok.Data;

@Data
class RowAttributes {
			
	public RowAttributes(int index, String name) {
	this.name = name;
	this.index = index;
	}
	private String name;
	private int index;
}
