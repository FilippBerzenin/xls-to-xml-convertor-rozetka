package xlsconvertor;

import java.util.Map;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.RequiredArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Item {
	
	private String ID;
	
	private String available;
	
	private String price_old;
	
	private String price;
	
	private String currencyId;
	
	private String categoryId;
	
	private int categoryIdNum;
	
	private String[] linksForPicture;
	
	private String stock_quantity;
	
	private String vendor;
	
	private String name;
	
	private String description;
	
	private Map<String, String> parameters;
	
}
