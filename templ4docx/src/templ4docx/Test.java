package templ4docx;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.variable.ImageType;
import pl.jsolve.templ4docx.variable.ImageVariable;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

public class Test {

	public static void main(String[] args) {
		Docx docx = new Docx("D:\\arcgis\\temp\\template.doc");

		Variables var = new Variables();
		var.addTextVariable(new TextVariable("${name}", "溧水区水库"));
		var.addTextVariable(new TextVariable("${date}", "2017年2月13日"));

		ImageVariable imageVariable1 = new ImageVariable("${photo1}",
				"D:\\arcgis\\temp\\test.jpg", ImageType.JPEG, 300, 300);
		ImageVariable imageVariable2 = new ImageVariable("${photo2}",
				"D:\\arcgis\\temp\\test.jpg", ImageType.JPEG, 150, 150);
		ImageVariable imageVariable3 = new ImageVariable("${photo3}",
				"D:\\arcgis\\temp\\test.jpg", ImageType.JPEG, 150, 150);
		var.addImageVariable(imageVariable1);
		var.addImageVariable(imageVariable2);
		var.addImageVariable(imageVariable3);
		docx.fillTemplate(var);
		docx.save("D:\\arcgis\\temp\\docs\\businessCard.doc");
	}

}
