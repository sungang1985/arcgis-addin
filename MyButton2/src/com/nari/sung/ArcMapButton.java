package com.nari.sung;

import java.io.IOException;

import javax.swing.JOptionPane;

import com.esri.arcgis.addins.desktop.Button;
import com.esri.arcgis.arcmapui.IMxDocument;
import com.esri.arcgis.carto.IActiveView;
import com.esri.arcgis.carto.IElement;
import com.esri.arcgis.carto.IFrameElement;
import com.esri.arcgis.carto.IGraphicsContainer;
import com.esri.arcgis.carto.IMap;
import com.esri.arcgis.carto.IMapFrame;
import com.esri.arcgis.carto.IMapSurround;
import com.esri.arcgis.carto.IMapSurroundFrame;
import com.esri.arcgis.carto.IMarkerNorthArrow;
import com.esri.arcgis.carto.IPageLayout;
import com.esri.arcgis.controls.MapControl;
import com.esri.arcgis.display.ICharacterMarkerSymbol;
import com.esri.arcgis.display.IMarkerSymbol;
import com.esri.arcgis.framework.IApplication;
import com.esri.arcgis.geometry.Envelope;
import com.esri.arcgis.geometry.IEnvelope;
import com.esri.arcgis.geometry.IPoint;
import com.esri.arcgis.geometry.Point;
import com.esri.arcgis.interop.AutomationException;
import com.esri.arcgis.output.ExportJPEG;
import com.esri.arcgis.output.IExport;
import com.esri.arcgis.system.IUID;
import com.esri.arcgis.system.UID;
import com.esri.arcgis.system.tagRECT;

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
			IPageLayout pageLayout = mxDocument.getPageLayout();

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
			export.cleanup();
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
