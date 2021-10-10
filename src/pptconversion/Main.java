package pptconversion;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Writer;
import java.lang.ref.WeakReference;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import javax.imageio.ImageIO;
import java.io.OutputStreamWriter;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.sl.draw.Drawable;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;



public class Main {
	/* Program returns exitcode 11 if successful, 0 if not. Takes dirs as args.
	 */
	public static void main(String[] args) throws IOException {
		
		if (args.length==0) {
			String path = Main.class.getProtectionDomain().getCodeSource().getLocation().getPath();
			String decodedPath = URLDecoder.decode(path, "UTF-8");
			File pathX = new File(decodedPath);
			String pdftotextPath = "";
			
			if (createDirs(pathX)) {
					createJSON(pathX, pdftotextPath);
			}
		} else {
				String pdftotextPath = args[1];
				createJSON(new File(args[0]), pdftotextPath);
			
		}
		System.out.println("Program terminated");
		System.exit(11);
			
	}
	
	public static Boolean createDirs(File dir) {
		Path path = Paths.get(dir+"\\data");
	
		if (!Files.exists(path)) {
			File f = new File(dir+"\\data");
			if (f.mkdir()) {
				System.out.println("Creating "+dir+"\\data"+" ..");
			} else {
			
				return false;
			}
		} else {
			
		}
		path = Paths.get(dir+"\\images");
		if (!Files.exists(path)) {
			File f = new File(dir+"\\images");
			
			if (f.mkdir()) {
				System.out.println("Creating "+dir+"\\images"+" ..");
			} else {
				
				return false;
			}
		} else {
			
		}
		
		
		
		
		return true;
	}
	
	public static void createJSON(File dir, String pdftotextpath) {
		System.out.println("Deconstructing files in " + dir + " ..");
		String write = "[";
		File[][] list = getPPTFiles(dir);
		if (list[0].length == 0 && list[1].length == 0 && list[2].length == 0) {
			System.out.println("No suitable files found.");
			System.exit(0);
		} else {
			createDirs(dir);
		}
		createTXTfromPDF(list[2], dir, pdftotextpath);
		splitPDF(list[2], dir);
		
		try {
			write+=ppt(list, dir);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		if (write.length() > 1) {
			write+=",";
		}
		write+=getPDFtextJSON(list[2], dir);
		write+="]";
		
			Writer myWriter = null;

			try {
				File filex = new File( dir+"\\data\\"+"data.json" );
				FileOutputStream fos = new FileOutputStream(filex);
				
				/*myWriter = new FileWriter(dir+"//data//"+"data.json");*/
				myWriter = new OutputStreamWriter(fos, StandardCharsets.UTF_8);
				BufferedWriter writer = new BufferedWriter(myWriter);
				writer.append(write);
				writer.close();
			
			/*	myWriter.write(write);
				myWriter.close(); */
			} catch (IOException e) {
				System.out.println("An error occurred.");
			      e.printStackTrace();
			}
			
		
	}
	
	public static Boolean splitPDF(File[] files, File dir) {
		
		for (int i=0; i<files.length;i++) {
			try {
				PDDocument document = PDDocument.load(new File(dir+"//"+files[i].getName()));
				PDFRenderer pdfRenderer = new PDFRenderer (document);
				for (int page = 0; page < document.getNumberOfPages(); ++page) {
					BufferedImage bim =  pdfRenderer.renderImageWithDPI(page, 300, ImageType.RGB);
					ImageIOUtil.writeImage(bim, dir+"//images//"+files[i].getName()+page+".png",300);
				}
				document.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
				return false;
			}
		}
		
		return true;
	}
	
	public static Boolean createTXTfromPDF(File[] files, File dir, String pdftotextPath) {
		
		for (int i = 0; i < files.length; i++) {
			
				try {
					String path = Main.class.getProtectionDomain().getCodeSource().getLocation().getPath();
					String decodedPath = URLDecoder.decode(path, "UTF-8");
					File pathX = new File(decodedPath);
					if (pdftotextPath.length() != 0) {
						pathX = new File(pdftotextPath);
					}
					Process process = new ProcessBuilder(pathX+"/pdftotext.exe",files[i]+"",dir+"/data/"+files[i].getName()+".txt").start();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
		}
		
		return true;
	}
		
	
	
	public static String getPDFtextJSON(File[] files, File dir) {
		if (files==null)return "";
		
		String result = "";
		
		for (int i=0;i<files.length;i++) {
			if (i>0) {
				result+=",";
			}
			result+=modifyPDFTextToJSON(dir+"/data/"+files[i].getName()+".txt");
		}
		return result;
	}
	
	public static String removeSpecialCharacters(String str) {
		
		return "";
	}

	public static String modifyPDFTextToJSON(String file) {
		String result="";
		try {	
			BufferedReader reader  = new BufferedReader(
				    new InputStreamReader(new FileInputStream(file),"ISO-8859-1"));
			result = "{\"file\": \"" + file + "\", \"pages\": [";
			String line;
			int index = 0;
			Boolean first = true;
			
			while ((line = reader.readLine()) != null) {	
				if (first) {
					result+="{\"id\": \"" + index + "\", \"title\": \"" + line + "\", \"text\": \"";
					first=false;
				}
				if (line.length()!=0 && line.charAt(0)=='\f') {
					index++;
					result+="\"},{\"id\": \""+index+"\", \"title\": \"" + line.substring(1) + "\", \"text\": \""+ line.substring(1) +" ";	
				} else {
					if (line.length()!=0) {
						String temp = line.replace("\"", "'");
						result+=" "+temp;
					} 
					
					
				}
				
			}
			reader.close();
			//System.out.println(modifyJSON(result));
		} catch (Exception e) {
			 System.err.format("Exception occurred trying to read '%s'.", file);
			  e.printStackTrace();
		}
		result=result.substring(0, result.length()-39);
		return modifyJSON(result+"\"}]}");
	}
	

	public static File[][] getPPTFiles(File dir){
			
			
			File[] filesPPT = dir.listFiles(new FilenameFilter() {
				@Override
				public boolean accept(File dir, String name) {
					return (name.toLowerCase().endsWith("ppt"));
				}
			});
			File[] filesPPTX = dir.listFiles(new FilenameFilter() {
				@Override
				public boolean accept(File dir, String name) {
					return (name.toLowerCase().endsWith("pptx"));
				}
			});
			File[] filesPDF = dir.listFiles(new FilenameFilter() {
				@Override
				public boolean accept(File dir, String name) {
					return (name.toLowerCase().endsWith("pdf"));
				}
			});
			File[][] files = new File[][] {filesPPT, filesPPTX, filesPDF};
			return files;
	}
	
	public static String modifyJSON(String str) {
		for (int i=0; i<str.length(); i++) {
			if (str.charAt(i) == '\\') {
				str = str.substring(0,i) + '\\' + str.substring(i);
				i++;
			}
		}
		for (int i=2; i<10; i++) {
			
			str = str.replace("  "," ");
		}
		
		return str.replace("\n", " ").replace("\r"," ").replace("\t", " ");
	}

	public static String ppt(File[][] args, File dir) throws IOException {
		
		File[] PPT = args[0];
		File[] PPTX = args[1];
		
		String result  = "";
		

		
		for (int i=0; i < PPTX.length; i++) {
			String paragraph = "";
			//String current = PPTX[i].getCanonicalPath().substring(0,PPTX[i].getCanonicalPath().length()-4);
			try {
				
				XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(PPTX[i]));
				Dimension pgsize = ppt.getPageSize();
				List<XSLFSlide> slides = ppt.getSlides();
				
			    BufferedImage img = null;
			    FileOutputStream out = null;
				paragraph+="{\"file\": \""+PPTX[i].getCanonicalPath() +"\", \"pages\": [";
			    for (int j = 0; j < slides.size(); j++) {
			    	paragraph+="{\"id\": \"" + j +"\", \"text\": \"";
			    	List<XSLFShape> shapes = slides.get(j).getShapes();
			    	for (XSLFShape shape: shapes) {
		        		if (shape instanceof XSLFTextShape) {
		        	        XSLFTextShape textShape = (XSLFTextShape)shape;
		        	        String text = textShape.getText();
		        	        paragraph+= text+" ";
		        		}
		        	}
			    	paragraph+="\"}";
			    	if (j!= slides.size()-1) {
			    		paragraph+=",";
			    	}
			    	
			    	
			    	img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
			    	Graphics2D graphics = img.createGraphics();
			    	
			        graphics.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
			        graphics.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
			        graphics.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);
			        graphics.setRenderingHint(RenderingHints.KEY_FRACTIONALMETRICS, RenderingHints.VALUE_FRACTIONALMETRICS_ON);
			        graphics.setRenderingHint(Drawable.BUFFERED_IMAGE, new WeakReference<>(img));
			    	
			        slides.get(j).draw(graphics);
					out = new FileOutputStream(dir+"/images/"+PPTX[i].getName()+j+".png");
			        ImageIO.write(img, "PNG", out);
			        graphics.dispose();
			        img.flush();
			        out.close();
			    }

			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			paragraph+="]}";
			result+=paragraph;
			
		}
		if (result.length()>1 && PPT.length!=0) {
			result+=",";
		}
		if (PPT.length>0) {
			for (int i=0; i < PPT.length; i++) {
				String paragraph = "";
				
				try {
					FileInputStream is;
					is = new FileInputStream(PPT[i]);
					HSLFSlideShow ppt = new HSLFSlideShow(is);
					is.close();
					
					Dimension pgsize = ppt.getPageSize();
					
					int idx = 0;
					paragraph+="{\"file\": \""+PPT[i].getCanonicalPath() +"\", \"pages\": [";
					for (HSLFSlide slide : ppt.getSlides()) {
						BufferedImage img = new BufferedImage(pgsize.width, pgsize.height, BufferedImage.TYPE_INT_RGB);
						Graphics2D graphics = img.createGraphics();
						
						graphics.setPaint(Color.white);
						graphics.fill(new Rectangle2D.Float(0,0, pgsize.width, pgsize.height));
						
						slide.draw(graphics);
						List<List<HSLFTextParagraph>> text = slide.getTextParagraphs();
					
						paragraph+="{\"id\": \"" + idx +"\", \"text\": \"";
						for (int j = 0; j < text.toArray().length; j++) {
							paragraph += text.toArray()[j];
						}
						paragraph+="\"}";
				    	if (idx!= ppt.getSlides().size()-1) {
				    		paragraph+=",";
				    	}
						
				    	
						FileOutputStream out = new FileOutputStream(dir+"/images/"+PPT[i].getName()+idx+".png");
						javax.imageio.ImageIO.write(img,  "png", out);
						out.close();
						idx++;
						}
					paragraph+="]}";
					result+=paragraph;
					ppt.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				//System.out.println(i + " "+paragraph);
			}
	}
		
		//System.out.println(modifyJSON(result));
		return(modifyJSON(result));
		
		
	}

}