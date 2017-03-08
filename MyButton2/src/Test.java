import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

public class Test {

	public static void main(String[] args) {
		try {
			File imageFile = new File("D:\\arcgis\\temp\\test.jpg");

			FileInputStream fis = new FileInputStream(imageFile);
			System.out.println(fis);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
}
