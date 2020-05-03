package org.libreoffice.example.comp;

import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XPropertyReplace;
import com.sun.star.util.XReplaceDescriptor;
import com.sun.star.util.XReplaceable;

import utils.*;

import com.sun.star.lib.uno.helper.Factory;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.FileSystem;
import java.util.ArrayList;
import java.util.Arrays;

import org.libreoffice.example.dialog.ActionOneDialog;
import org.libreoffice.example.helper.DialogHelper;

import com.sun.star.awt.FontWeight;
import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.text.XTextDocument;
import com.sun.star.lib.uno.helper.WeakBase;

import opennlp.tools.sentdetect.SentenceDetector;
import opennlp.tools.sentdetect.SentenceDetectorME;
import opennlp.tools.sentdetect.SentenceModel;
import opennlp.tools.util.InvalidFormatException;

public final class KlaroImpl extends WeakBase
   implements com.sun.star.lang.XServiceInfo,
              com.sun.star.task.XJobExecutor
{
    private final XComponentContext m_xContext;
    private static final String m_implementationName = KlaroImpl.class.getName();
    private static final String[] m_serviceNames = {
        "org.libreoffice.example.StarterProject" };
    private XComponentContext xcc;
	private XTextDocument textDoc;    


    public KlaroImpl( XComponentContext context )
    {
        m_xContext = context;
    };

    public static XSingleComponentFactory __getComponentFactory( String sImplementationName ) {
        XSingleComponentFactory xFactory = null;

        if ( sImplementationName.equals( m_implementationName ) )
            xFactory = Factory.createComponentFactory(KlaroImpl.class, m_serviceNames);
        return xFactory;
    }

    public static boolean __writeRegistryServiceInfo( XRegistryKey xRegistryKey ) {
        return Factory.writeRegistryServiceInfo(m_implementationName,
                                                m_serviceNames,
                                                xRegistryKey);
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
         return m_implementationName;
    }

    public boolean supportsService( String sService ) {
        int len = m_serviceNames.length;

        for( int i=0; i < len; i++) {
            if (sService.equals(m_serviceNames[i]))
                return true;
        }
        return false;
    }

    public String[] getSupportedServiceNames() {
        return m_serviceNames;
    }

    // com.sun.star.task.XJobExecutor:
    public void trigger(String action)
    {
    	switch (action) {
    	case "actionOne":
    		ActionOneDialog actionOneDialog = new ActionOneDialog(m_xContext);
    		actionOneDialog.show();
    		xcc = m_xContext;
    		XComponent doc = Lo.addonInitialize(xcc);
    		textDoc = Write.getTextDoc(doc);
			worker("koniec");
    		break;
    	default:
    		DialogHelper.showErrorMessage(m_xContext, null, "Unknown action: " + action);
    	}
        
    }
    
   
    
    private int worker(String searchKey) 
    /* Only matches whole words and case sensitive, and highlight
       in bold and red;  **ADDED** this function */
    {
      System.out.println("look for: " + searchKey);
      String s = textDoc.getText().getString();
     //InputStream modelIn = null;
      String[] sentences = null;
      SentenceModel model;
	try {
		model = new SentenceModel(new File("/en-sent.bin"));
	    SentenceDetector sentenceDetector = new SentenceDetectorME(model);
	    System.out.println(s);
	    sentences = sentenceDetector.sentDetect(s);
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
    
      //Printing the sentences 
      for(String sent : sentences) {        
         System.out.println(sent); 
      }

      XReplaceable repl = Lo.qi(XReplaceable.class, textDoc);
      XReplaceDescriptor desc = repl.createReplaceDescriptor();

      // Gets a XPropertyReplace object for altering the properties
      // of the replaced text
      XPropertyReplace propReplace = Lo.qi(XPropertyReplace.class, desc);

      // Set the replaced text to bold and red
      PropertyValue wv = new PropertyValue("CharWeight", -1, FontWeight.BOLD,
                                                      PropertyState.DIRECT_VALUE);
      PropertyValue cv = new PropertyValue("CharColor", -1, Color.RED.getRGB(),
                                                     PropertyState.DIRECT_VALUE);
      PropertyValue[] props = new PropertyValue[] {cv, wv};
      try {
        propReplace.setReplaceAttributes(props);

        // Only match whole words and case sensitive
        desc.setPropertyValue("SearchCaseSensitive", true);
        desc.setPropertyValue("SearchWords", true);
      }
      catch (com.sun.star.uno.Exception ex) {
        System.out.println("Error setting up search properties");
        return -1;
      }

      // Replaces all instances of searchKey with new Text properties
      // and gets the number of instances of the searchKey
      desc.setSearchString(searchKey);
      desc.setReplaceString(searchKey);
      return repl.replaceAll(desc);
    }  // end of applyEzHighlighting()
    

}
