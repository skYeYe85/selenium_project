package selenium.webdriver;

public class Einkommen {
	
	private int brutto;
	private double jbrutto;
	private double netto;
	private double jnetto;
	
	public Einkommen(){
		
	}
	
	public Einkommen(int brutto, double jbrutto, double netto, double jnetto){
		this.brutto = brutto;
		this.jbrutto = jbrutto;
		this.netto = netto;
		this.jnetto = jnetto;
	}

	public double getBrutto() {
		return brutto;
	}

	public void setBrutto(int brutto) {
		this.brutto = brutto;
	}

	public double getJbrutto() {
		return jbrutto;
	}

	public void setJbrutto(double jbrutto) {
		this.jbrutto = jbrutto;
	}

	public double getNetto() {
		return netto;
	}

	public void setNetto(double netto) {
		this.netto = netto;
	}

	public double getJnetto() {
		return jnetto;
	}

	public void setJnetto(double jnetto) {
		this.jnetto = jnetto;
	}
}
