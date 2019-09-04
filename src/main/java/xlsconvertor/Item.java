package xlsconvertor;

import java.util.Map;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class Item {
	
	private int ID;
	
	private String available;
	
	private String price_old;
	
	private String price;
	
	private String currencyId;
	
	private String categoryId;
	
	private String[] linksForPicture;
	
	private String stock_quantity;
	
	private String vendor;
	
	private String name;
	
	private String description;
	
	private Map<String, String> parameters;

}
