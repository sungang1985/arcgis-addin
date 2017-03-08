import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import com.healthmarketscience.jackcess.DatabaseBuilder;
import com.healthmarketscience.jackcess.Row;
import com.healthmarketscience.jackcess.Table;

public class TestJackcess {

	public static void main(String[] args) {
		/*try {
			Table table = DatabaseBuilder.open(
					new File("D:\\arcgis\\temp\\iglp.mdb")).getTable("IGLP_P");
			for (Row row : table) {
				System.out.println("Column 'a' has value: "
						+ new BigDecimal(row.get("HZB").toString())
								.toPlainString());
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}*/
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
		System.out.println(sdf.format(new Date()));

		System.out.println(String.format("%04d", 1));

	}
}
