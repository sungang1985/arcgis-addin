package com.nari.sung;

import java.io.IOException;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.variable.ImageType;
import pl.jsolve.templ4docx.variable.ImageVariable;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

import com.esri.arcgis.addins.desktop.Button;
import com.esri.arcgis.arcmapui.IMxDocument;
import com.esri.arcgis.carto.IFeatureLayer;
import com.esri.arcgis.carto.IMap;
import com.esri.arcgis.framework.IApplication;
import com.esri.arcgis.geodatabase.IFeature;
import com.esri.arcgis.geodatabase.IFeatureClass;
import com.esri.arcgis.geodatabase.IFeatureCursor;
import com.esri.arcgis.interop.AutomationException;

public class ArcMapButton extends Button {
	private IApplication application;

	// private IMxDocument mxDocument;
	// private IMap map;
	// private IPageLayout pageLayout;

	/**
	 * Called when the button is clicked.
	 * 
	 * @exception java.io.IOException
	 *                if there are interop problems.
	 * @exception com.esri.arcgis.interop.AutomationException
	 *                if the component throws an ArcObjects exception.
	 */
	@Override
	public void onClick() throws IOException, AutomationException {
		try {
			IMxDocument mxDocument = (IMxDocument) application.getDocument();
			IMap map = mxDocument.getFocusMap();

			IFeatureLayer featureLayer = (IFeatureLayer) map.getLayer(0);
			IFeatureClass featureClass = featureLayer.getFeatureClass();
			IFeatureCursor updateCursor = featureClass.search(null, false);
			int fidFieldIndex = featureClass.findField("AB");
			int xhFieldIndex = featureClass.findField("XH");
			int hzbFieldIndex = featureClass.findField("HZB");
			int zzbFieldIndex = featureClass.findField("ZZB");
			int gcFieldIndex = featureClass.findField("GC");
			IFeature feature = null;

			Docx docx = new Docx("D:\\arcgis\\temp\\template.docx");
			while ((feature = updateCursor.nextFeature()) != null) {

				Variables var = new Variables();
				var.addTextVariable(new TextVariable("${name}", "溧水区水库"));
				var.addTextVariable(new TextVariable("${date}", "2017年3月4日"));

				ImageVariable imageVariable1 = new ImageVariable("${photo1}",
						"D:\\arcgis\\temp\\界桩身份证图片\\L0001J.JPG",
						ImageType.JPEG, 300, 300);
				ImageVariable imageVariable2 = new ImageVariable("${photo2}",
						"D:\\arcgis\\temp\\界桩身份证图片\\L0001J.JPG",
						ImageType.JPEG, 150, 150);
				ImageVariable imageVariable3 = new ImageVariable("${photo3}",
						"D:\\arcgis\\temp\\界桩身份证图片\\L0001Y.JPG",
						ImageType.JPEG, 150, 150);
				var.addImageVariable(imageVariable1);
				var.addImageVariable(imageVariable2);
				var.addImageVariable(imageVariable3);
				docx.fillTemplate(var);

				docx.save("D:\\arcgis\\temp\\界桩身份证文档\\"
						+ String.valueOf(feature.getValue(xhFieldIndex))
						+ ".docx");
			}

			/*IPageLayout pageLayout = mxDocument.getPageLayout();

			IActiveView activeView = mxDocument.getActiveView();

			IPoint point = new Point();
			point.putCoords(40536325.385, 3743761);

			IEnvelope pCurrentEnvelop = activeView.getExtent();
			pCurrentEnvelop.centerAt(point);
			pCurrentEnvelop.expand(0.125, 0.125, true);
			activeView.setExtent(pCurrentEnvelop);
			activeView.refresh();

			pCurrentEnvelop = new Envelope();
			pCurrentEnvelop.putCoords(2, 25, 2, 25);

			IUID uid = new UID();
			uid.setValue("esriCarto.MarkerNorthArrow");
			IGraphicsContainer graphicsContainer = (IGraphicsContainer) pageLayout;
			activeView = (IActiveView) pageLayout;
			IFrameElement frameElement = graphicsContainer.findFrame(map);
			IMapFrame mapFrame = (IMapFrame) frameElement;
			IMapSurroundFrame mapSurroundFrame = mapFrame.createSurroundFrame(
					uid, null);
			IElement element = (IElement) mapSurroundFrame;
			element.setGeometry(pCurrentEnvelop);
			element.activate(activeView.getScreenDisplay());
			graphicsContainer.addElement(element, 0);

			IMapSurround mapSurround = mapSurroundFrame.getMapSurround();

			IMarkerNorthArrow markerNorthArrow = (IMarkerNorthArrow) mapSurround;
			IMarkerSymbol markerSymbol = markerNorthArrow.getMarkerSymbol();
			ICharacterMarkerSymbol characterMarkerSymbol = (ICharacterMarkerSymbol) markerSymbol;
			characterMarkerSymbol.setCharacterIndex(200);
			markerNorthArrow.setMarkerSymbol(characterMarkerSymbol);

			IExport export = new ExportJPEG();
			export.setExportFileName("D:\\arcgis\\temp\\test.jpg");
			export.setResolution(96);

			tagRECT exportRECT = activeView.getExportFrame();
			IEnvelope envelope = new Envelope();

			envelope.putCoords(exportRECT.left, exportRECT.top,
					exportRECT.right, exportRECT.bottom);
			export.setPixelBounds(envelope);

			int hDC = export.startExporting();
			activeView.output(hDC, (int) export.getResolution(), exportRECT,
					null, null);

			export.finishExporting();
			export.cleanup();*/
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@Override
	public void init(IApplication application) throws IOException,
			AutomationException {
		this.application = application;
	}

}
