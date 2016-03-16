package in.setia.document.process;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.MissingResourceException;
import java.util.ResourceBundle;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class TextExtractor {
	
	private static final ResourceBundle config;
    private static OfficeManager connection = null;

    
	static {
		config = ResourceBundle.getBundle("config");
	}
	
	public String extractTextFromFile(String inputFile) throws DocumentConversionException {
		validateParameters(inputFile);
		String outputFile = inputFile.replaceAll("doc", "txt");
		return extract(inputFile, outputFile);
	}

	public String extractHtmlTextFromFile(String inputFile) throws DocumentConversionException {
		validateParameters(inputFile);
		String outputFile = inputFile.replaceAll("doc", "html");
     	String outputFileRtf = null;
		if(inputFile.toLowerCase().endsWith("doc")){
			outputFileRtf = inputFile.replaceAll("doc", "rtf");
     		extract(inputFile, outputFileRtf);
     	} else {
     		outputFileRtf = outputFile;
     	}
     	return extract(outputFileRtf, outputFile);
	}
	
	private String extract(String inputFile, String outputFile) throws DocumentConversionException {
		OfficeDocumentConverter converter = new OfficeDocumentConverter(getConnection());  	
		converter.convert(new File(inputFile), new File(outputFile));
		BufferedReader br = null;
		try {
			br = new BufferedReader(new FileReader(outputFile));
			String var = br.readLine();
			StringBuffer sb = new StringBuffer();
			while( var != null){
				sb.append(var);
			}
			return sb.toString();
		} catch (IOException e) {
			throw new DocumentConversionException("Error reading output", e);
		} finally {
			try {
				br.close();
			} catch (IOException e) {
				System.out.println("error while closing "+e.getMessage());
			}
			br = null;
		}
		
	}
	
	public void generateRtfDocumentFromHtml(String inputFile,String outputFile) throws DocumentConversionException {
		OfficeDocumentConverter converter = new OfficeDocumentConverter(getConnection()); 
     	System.out.println("OO connection initialized...");
		converter.convert(new File(inputFile), new File(outputFile));
	}
	
	//	@SuppressWarnings("unchecked")
	public void convertToPDF(String inputFile, String outputFile, String pwd) throws IOException, DocumentConversionException {
		// create a PDF DocumentFormat (as normally configured in document-formats.xml)
		DocumentFormat customPdfFormat = new DocumentFormat("Portable Document DateFormatProcesser", "application/pdf", "pdf");
//		customPdfFormat.setExportFilter(DocumentFamily.TEXT, "writer_pdf_Export");
		System.out.println("Password is "+pwd);
		// now set our custom options
		Map<String, Object> pdfOptions = new HashMap<String, Object>();
		pdfOptions.put("EncryptFile", Boolean.TRUE);
		if(pwd!=null)
			pdfOptions.put("DocumentOpenPassword", pwd);
//		customPdfFormat.setExportOption(DocumentFamily.TEXT, "FilterData", pdfOptions);
		//TODO to enable setting pdf options with JODConverter 3.0

		OfficeDocumentConverter converter = new OfficeDocumentConverter(getConnection());  	
		converter.convert(new File(inputFile), new File(outputFile), customPdfFormat);
    }


//	public void convertDocumentToPDF(String inputFile, String outputFile, String pwd) throws IOException {
//		// create a PDF DocumentFormat (as normally configured in document-formats.xml)
//		DocumentFormat customPdfFormat = new DocumentFormat("Portable Document Format", "application/pdf", "pdf");
//		customPdfFormat.setExportFilter(DocumentFamily.TEXT, "writer_pdf_Export");
//		System.out.println("Password is "+pwd);
//		// now set our custom options
//		Map pdfOptions = new HashMap();
//		pdfOptions.put("EncryptFile", Boolean.TRUE);
//		if(pwd!=null)
//			pdfOptions.put("DocumentOpenPassword", pwd);
//		customPdfFormat.setExportOption(DocumentFamily.TEXT, "FilterData", pdfOptions);
//
//
//     	DocumentConverter converter = new OpenOfficeDocumentConverter(getConnection());  	
//		converter.convert(new File(inputFile), new File(outputFile));
//    }

	
	public static void stopOpenOfficeProcess(){
		connection.stop();
	}
	
	private static OfficeManager getConnection() throws DocumentConversionException {
		try {
			if (connection == null) {
				connection = new DefaultOfficeManagerConfiguration()
			      				.setOfficeHome(config.getString("office.home"))
			      				.setPortNumber(Integer.parseInt(config.getString("office.port")))
			      				.setTaskExecutionTimeout(15000L)
			      				.setMaxTasksPerProcess(10000)
			      				.buildOfficeManager();
				connection.start();
			}
		} catch (MissingResourceException e) {
			throw new DocumentConversionException("Config file missing for Open Office");
		}
		return connection;
	}
		
	private static void validateParameters(String inputFile)
			throws DocumentConversionException {
		if(inputFile==null || "".equals(inputFile))
			throw new DocumentConversionException("No input file");
		if(!inputFile.toLowerCase().endsWith("docx") && !inputFile.toLowerCase().endsWith("doc") && !inputFile.toLowerCase().endsWith("rtf"))
			throw new DocumentConversionException("Input document format not supported :"+inputFile);
		if(!inputFile.toLowerCase().endsWith("docx") && !inputFile.toLowerCase().endsWith("doc") && !inputFile.toLowerCase().endsWith("rtf"))
			throw new DocumentConversionException("Input document format not supported");
		if(!inputFile.toLowerCase().endsWith("htm") && !inputFile.toLowerCase().endsWith("html"))
			throw new DocumentConversionException("Input document format not supported");
		if(!inputFile.toLowerCase().endsWith("htm") && !inputFile.toLowerCase().endsWith("html"))
			throw new DocumentConversionException("Input document format not supported");
	}

	@Override
	protected void finalize() throws Throwable {
		stopOpenOfficeProcess();
	}
	


}
