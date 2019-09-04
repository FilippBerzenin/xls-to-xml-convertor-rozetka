package xlsconvertor;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class Item {
	
	private int ID;
	
	private String available;
	
	private double price_old;
	
	private double price;
	
	private String currencyId;
	
	private String categoryId;
	
	private String[] linksForPicture;
	
	private int stock_quantity;
	
	private String vendor;
	
	private String name;
	
	private String description;
	
	private String[] parameters;

}
