package in.setia.document.process;

import java.io.File;
import java.util.MissingResourceException;
import java.util.ResourceBundle;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

public class TextExtractor {
	
	private static final ResourceBundle config;
    private static OfficeManager connection = null;

    
	static {
		config = ResourceBundle.getBundle("resources.config");
	}

	public void extractTextFromDocument(String inputFile,String outputFile) throws DocumentConversionException {
		validateParameters(inputFile, outputFile);
		if(!inputFile.toLowerCase().endsWith("docx") && !inputFile.toLowerCase().endsWith("doc") && !inputFile.toLowerCase().endsWith("rtf"))
			throw new DocumentConversionException("Input document format not supported :"+inputFile);
		if(!outputFile.toLowerCase().endsWith("txt"))
			throw new DocumentConversionException("Output file for this method can only be a valid text file :"+outputFile);
		OfficeDocumentConverter converter = new OfficeDocumentConverter(getConnection());  	
		converter.convert(new File(inputFile), new File(outputFile));
	}
	
	public void extractHtmlFromDocument(String inputFile, String outputFile) throws DocumentConversionException {
		validateParameters(inputFile, outputFile);
		if(!inputFile.toLowerCase().endsWith("docx") && !inputFile.toLowerCase().endsWith("doc") && !inputFile.toLowerCase().endsWith("rtf"))
			throw new DocumentConversionException("Input document format not supported");
		if(!outputFile.toLowerCase().endsWith("htm") && !outputFile.toLowerCase().endsWith("html"))
			throw new DocumentConversionException("Output file for this method can only be a valid HTML file");
		OfficeDocumentConverter converter = new OfficeDocumentConverter(getConnection()); 
     	if(inputFile.toLowerCase().endsWith("doc")){
     		converter.convert(new File(inputFile), new File(inputFile.replaceAll("doc", "rtf")));
     		inputFile = inputFile.replaceAll("doc", "rtf");
     	}
		converter.convert(new File(inputFile), new File(outputFile));		
	}
	
	public void extractTextFromHtml(String inputFile,String outputFile) throws DocumentConversionException {
		if(!inputFile.toLowerCase().endsWith("htm") && !inputFile.toLowerCase().endsWith("html"))
			throw new DocumentConversionException("Input document format not supported");
		if(!outputFile.toLowerCase().endsWith("txt"))
			throw new DocumentConversionException("Output file for this method can only be a valid text file");
		OfficeDocumentConverter converter = new OfficeDocumentConverter(getConnection()); 
     	System.out.println("OO connection initialized...");
		converter.convert(new File(inputFile), new File(outputFile));
	}

	public void generateRtfDocumentFromHtml(String inputFile,String outputFile) throws DocumentConversionException {
		if(!inputFile.toLowerCase().endsWith("htm") && !inputFile.toLowerCase().endsWith("html"))
			throw new DocumentConversionException("Input document format not supported");
		if(!outputFile.toLowerCase().endsWith("rtf"))
			throw new DocumentConversionException("Output file for this method can only be a valid RTF file");
		OfficeDocumentConverter converter = new OfficeDocumentConverter(getConnection()); 
     	System.out.println("OO connection initialized...");
		converter.convert(new File(inputFile), new File(outputFile));
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
	
	
	/**
	 * @param inputFile
	 * @param outputFile
	 * @throws DocumentConversionException
	 */
	private static void validateParameters(String inputFile, String outputFile)
			throws DocumentConversionException {
		if(inputFile==null || "".equals(inputFile))
			throw new DocumentConversionException("No input file");
		if(outputFile == null || "".equals(outputFile))
			throw new DocumentConversionException("No output file");
	}

	@Override
	protected void finalize() throws Throwable {
		stopOpenOfficeProcess();
	}
	


}
