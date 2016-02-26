package in.setia.document.process;

import java.util.ResourceBundle;

import ooo.connector.BootstrapSocketConnector;

import com.sun.star.beans.PropertyValue;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XReplaceDescriptor;
import com.sun.star.util.XReplaceable;

public class DocumentEditor {

	private static final ResourceBundle config;
	private static XComponentLoader xComponentLoader;

    
	static {
		config = ResourceBundle.getBundle("resources.config");
			try {
				XComponentContext xContext = BootstrapSocketConnector.bootstrap(config.getString("office.home"));
				XMultiComponentFactory xMCF = xContext.getServiceManager();
				Object desktop = xMCF.createInstanceWithContext("com.sun.star.frame.Desktop", xContext);
				xComponentLoader = (XComponentLoader)UnoRuntime.queryInterface(XComponentLoader.class, desktop);
			} catch (BootstrapException e) {
				e.getCause().printStackTrace();
				e.printStackTrace();
			} catch (Exception e) {
				e.printStackTrace();
			}
	}


	public void maskDocument(String originalDocument , String targetDocument , String[] sensitiveWords) throws DocumentConversionException {
		try {
			if(originalDocument == null || "".equalsIgnoreCase(originalDocument))
				throw new DocumentConversionException("Source Document cannot be empty.");
			if(targetDocument == null || "".equalsIgnoreCase(targetDocument))
				throw new DocumentConversionException("Target Document cannot be empty.");
			if (!targetDocument.endsWith("doc"))
				throw new DocumentConversionException("Target Document can only have a DOC extension.");
				
			originalDocument = "file://" + originalDocument;
			targetDocument = "file://" + targetDocument;

			PropertyValue[] loadProps = new PropertyValue[1];
			loadProps[0] = new PropertyValue();
			loadProps[0].Name="Hidden";     
			loadProps[0].Value = true;
			XComponent activeDoc = xComponentLoader.loadComponentFromURL(originalDocument , "_blank", 0, loadProps);

			//Replace contents
			XReplaceable xReplaceable = (XReplaceable)UnoRuntime.queryInterface(XReplaceable.class, activeDoc);
			XReplaceDescriptor xReplaceDescr = (XReplaceDescriptor)xReplaceable.createReplaceDescriptor();
			for (int i = 0; i < sensitiveWords.length; i++) {
				replaceContent(xReplaceDescr, xReplaceable, sensitiveWords[i]);
			}
			System.out.println("Content Replaced");
	        
			// Save File
			XStorable xStorable =(XStorable)UnoRuntime.queryInterface(XStorable.class, activeDoc );    
			loadProps = new PropertyValue[ 2 ];	 
			loadProps[0] = new PropertyValue();     
			loadProps[0].Name = "Overwrite";
			loadProps[0].Value = new Boolean(true);
			loadProps[1] = new PropertyValue();
			loadProps[1].Name = "FilterName";
			loadProps[1].Value = "MS Word 97";
			xStorable.storeAsURL(targetDocument ,loadProps );
			System.out.println("Saved.");
		} catch (Exception e) {
			e.printStackTrace();
			throw new DocumentConversionException(e);
		}  

	}
	
	private static void replaceContent(XReplaceDescriptor xReplaceDescr, XReplaceable xReplaceable, String searchString){
		if(searchString != null && !"".equalsIgnoreCase(searchString)){
			xReplaceDescr.setSearchString(searchString);
			String replaceString = "";
			for(int i=0;i<searchString.length();i++){        
				replaceString="*".concat(replaceString);
			}	              
			
			xReplaceDescr.setReplaceString(replaceString);     
			xReplaceable.replaceAll(xReplaceDescr);
		}
	}


}
