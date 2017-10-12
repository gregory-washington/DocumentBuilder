import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Set;

import org.apache.commons.io.FileUtils;



public class TestStringReplace {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		 try {
			 String filename = "C:\\temp\\MasterAgreementConversion.doc";
				
				HashMap <String, String> replacementText = new HashMap<String, String>();
				
				
				
					replacementText.put("[#masterNumber#]","MP05101");
					replacementText.put("[#amendmentNumberNumeric#]","0");
					replacementText.put("[#contractGenerationDate#]","12/22/2016"); 
					replacementText.put("[#contactName#]","Deborah Momsen-Hudsen"); 
					replacementText.put("[#contactTitle#]","Vice President & Director of Secondary Marketing"); 
					replacementText.put("[#addressee#]","Ms. Momsen-Hudsen"); 
					replacementText.put("[#lenderName#]","Self-Help Ventures Fund"); 
					replacementText.put("[#lenderNameCAPS#]","SELF-HELP VENTURES FUND"); 
					replacementText.put("[#mstrAffiliateNos#]"," FICR 31884-000-8 FICR 31885-000-3"); 
					replacementText.put("[#subjectMstrAffiliateNos#]","31884-000-8 FICR 31885-000-3"); 
					replacementText.put("[#mstrAffiliateNames#]",", Self Help Ventures  Fund, and Self Help  Ventures Fund"); 
					replacementText.put("[#mstrAffiliateNamesCR#]"," FICR Self Help Ventures Fund FICR Self Help Ventures Fund");
					replacementText.put("[#mstrAffiliateNamesCAPS#]","SELF HELP VENTURES FUND FIST# SELF HELP VENTURES FUND FIST# ");
					replacementText.put("[#agreementMandatoryAmount#]","");
					replacementText.put("[#agreementOptionalAmount#]","$950,000,000.00 (Optional)");
					replacementText.put("[#currentDate#]","December 22, 2016");
					replacementText.put("[#conversionExpirationDate#]","");
					replacementText.put("[#conversionEffectiveDate#]","");
					replacementText.put("[#streetAddress#]","301 West Main Street");
					replacementText.put("[#cityName#]","Durham");
					replacementText.put("[#stateCode#]","NC");
					replacementText.put("[#postalCode#]","27701");
					replacementText.put("[#lenderEmail#]","deborah.momsen-hudson@self-help.org;Brenda.farlow@self-help.org");
					replacementText.put("[#contractSignerName#]","C. Edmund Hohmann Jr.");
					replacementText.put("[#contractSignerTitleName#]","Vice President");
					replacementText.put("[#deliveryTerm_3#]","FDChar");
					replacementText.put("[#deliveryTerm_2#]","FDTR");
					replacementText.put("[#deliveryTerm#]",":");
					replacementText.put("[#deliveryTermNumeric#]","");
					replacementText.put("[#effectiveDate#]","January 1, 2017");
					replacementText.put("[#effectiveDateCurrent#]","January 1, 2017");
					replacementText.put("[#expirationDate#]","December 31, 2017");
					replacementText.put("[#lenderID#]","23733-000-5");
					replacementText.put("[#approvedDeliveryAmount#]","$950,000,000.00");
					replacementText.put("[#amendmentNumber#]","");
					replacementText.put("[#agreementTotalAmount#]","$950,000,000.00");
					replacementText.put("[#approvedTotalAmount#]","$950,000,000.00");
					replacementText.put("[#conversionAmount#]","");
					replacementText.put("[#estimatedAmount#]","FDTR");
					replacementText.put("[#commitTerm#]","12");
					replacementText.put("[#approvedMandatoryAmount#]","$0.00");
					replacementText.put("[#buyoutFee#]","FDTR");
					replacementText.put("[#buyoutFeeMaster#]","12.5");
					replacementText.put("[#riskBasedPricing#]","FDTR");
					replacementText.put("[#capitalConversionCycle#]","");
					replacementText.put("[#tolerance#]","");
					replacementText.put("[#cAMName#]","Kathy Mac");
					replacementText.put("[#fannieMaeAddress#]","One South Wacker Drive, Suite 1400, Chicago, IL  60606");
					replacementText.put("[#poolPurchaseContracts#]","N/A");
					replacementText.put("[#masterAmendmentKind#]","section below");
					replacementText.put("[#agreementMandatoryAmendmentLetter#]","$0.00");
					replacementText.put("[#expirationDateAmendmentLetter#]","December 31, 2017");
					replacementText.put("[#buyoutFeeAmendmentLetter#]","12.5");
					replacementText.put("[#toleranceAmendmentLetter#]","5%");
					replacementText.put("[#deliveryTermAmendmentLetter#]","1");
					replacementText.put("[#masterAmendmentConversionKind#]","FDTR");
					replacementText.put("[#varTOCDisplay#]","");
					replacementText.put("[#srTOCDisplay#]","FDTR");
					replacementText.put("[#varianceRenewedDisplay#]","FDTR");
					replacementText.put("[#varianceAddedDisplay#]","FDTR");
					replacementText.put("[#varianceDiscontinuedDisplay#]","FDTR");
					replacementText.put("[#srRenewedDisplay#]","FDTR");
					replacementText.put("[#srAddedDisplay#]","FDTR");
					replacementText.put("[#srDiscontinuedDisplay#]","FDTR");
					replacementText.put("[#mBSRenewedDisplay#]","FDTR");
					replacementText.put("[#mBSAddedDisplay#]","FDTR");
					replacementText.put("[#areBookmarksIndented#]","1");
					replacementText.put("[#cgType#]","Full");
					replacementText.put("[#isConversion#]","0");
					replacementText.put("[#masterKind#]","New Master Agreement");
				
				
			 
			 
			 
			 
		     String content = FileUtils.readFileToString(new File(filename), "UTF-8");
		     
		 
		     Set<String> keys = replacementText.keySet();
		     
		     for(String key : keys) {
		    	 content = content.replace(key, replacementText.get(key)); 
		     }
		     
		     
		     
		     
		     File tempFile = new File("C:\\temp\\copy.rtf");
		     FileUtils.writeStringToFile(tempFile, content, "UTF-8");
		  } catch (IOException e) {
		     //Simple exception handling, replace with what's necessary for your use case!
			  e.printStackTrace();
		     throw new RuntimeException("Generating file failed", e);
		  }

	}

}
