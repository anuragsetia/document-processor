package in.setia.document.process;

public class DocumentConversionException extends Exception {

	private static final long serialVersionUID = 1L;
	
	public DocumentConversionException(String msg){
		super(msg);
	}
	
	public DocumentConversionException(Exception ex){
		super(ex);
	}

}
