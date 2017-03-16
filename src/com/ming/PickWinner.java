package com.ming;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class PickWinner {
	static String dataFile="";
	String[] weightCategories = {"KenPom Rank","Def Eff","Off Eff","SoS","Free Throw","Off Reb","Turnover Margin","Star Player","Great Coach","Tournament Rank"};
	static Hashtable<String, TeamInfo> datahs = new Hashtable<String, TeamInfo>();
	static Hashtable<String,WeightInfo> weightHs = new Hashtable<String, WeightInfo>();
	static Hashtable<Integer,String> weightListHs	= new Hashtable<Integer, String>();
	static TreeMap<Integer,String> teamRanks = new TreeMap<Integer,String>();//kenpomRank, teamName
	static Hashtable<Integer,Integer> ranksList = new Hashtable<Integer,Integer>();//rank order, rank number
	static int hkenpomRank=0;
	static int htournamentRank=0;
	static double hdefEff=0;
	static double ldefEff=100000;	
	static double hoffEff=0;
	static double loffEff=100000;
	static double hsos=0;
	static double lsos=1000000;
	static double hfreeThrow=0;
	static double lfreeThrow=1000000;
	static double hoffReb=0;
	static double loffReb=1000000;
	static double hturnoverMargin=0;
	static double lturnoverMargin=1000000;
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		getDataFile();
		if (dataFile==""){
			System.out.println("There is not data to process.");
			return;
		}else{
			System.out.println("There is data file to use>>>" + dataFile);
		}
		try {
			getInputData();
			calculateWeight();
			pickAWinner();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	private static void pickAWinner() {
	      Set<Entry<Integer, String>> set = teamRanks.entrySet();
	      // Get an iterator
	      Iterator i = set.iterator();
	      int j=0;
	      // Display elements
	      while(i.hasNext()) {
	    	  j++;
	        Map.Entry<Integer, String> me = (Map.Entry<Integer, String>)i.next();
	        ranksList.put(j, me.getKey());
	      }		
	      //first round
	      int totTeam = teamRanks.size();
	      TreeMap<Integer, String> eastSide = new TreeMap<Integer, String>();
	      TreeMap<Integer, String> westSide = new TreeMap<Integer, String>();
	      
	      for (int teamIdx=1;teamIdx<=totTeam/4;teamIdx++){
	    	  int firstTeamRank = ranksList.get(teamIdx);
	    	  int secondTeamRank = ranksList.get(totTeam -teamIdx -1);
	    	  String firstTeamName = teamRanks.get(firstTeamRank);
	    	  String secondTeamName = teamRanks.get(secondTeamRank);
	    	  TeamInfo firstTeamInfo = datahs.get(firstTeamName);
	    	  TeamInfo secondTeamInfo = datahs.get(secondTeamName);
	    	  if (firstTeamInfo.totalScore > secondTeamInfo.totalScore){
	    		  eastSide.put(teamIdx, firstTeamName + " recent rank: " + firstTeamRank + " total score is: " + firstTeamInfo.totalScore + ". (Loser: " + secondTeamName +" with scoure " + secondTeamInfo.totalScore + ")");
	    	  }else{
	    		  eastSide.put(totTeam -teamIdx -1, secondTeamName + " recent rank: " + secondTeamRank + " total score is: " + secondTeamInfo.totalScore + ". (Loser: " + firstTeamInfo +" with scoure " + firstTeamInfo.totalScore + ")");
	    	  }
	    	  firstTeamRank = ranksList.get(teamIdx+1);
	    	  secondTeamRank = ranksList.get(totTeam -teamIdx -2);	    	  
	    	  firstTeamName = teamRanks.get(firstTeamRank);
	    	  secondTeamName = teamRanks.get(secondTeamRank);
	    	  firstTeamInfo = datahs.get(firstTeamName);
	    	  secondTeamInfo = datahs.get(secondTeamName);
	    	  if (firstTeamInfo.totalScore > secondTeamInfo.totalScore){
	    		  westSide.put(teamIdx+1, firstTeamName + " recent rank: " + firstTeamRank + " total score is: " + firstTeamInfo.totalScore + ". (Loser: " + secondTeamName +" with scoure " + secondTeamInfo.totalScore + ")");
	    	  }else{
	    		  westSide.put(totTeam -teamIdx -2, secondTeamName + " recent rank: " + secondTeamRank + " total score is: " + secondTeamInfo.totalScore + ". (Loser: " + firstTeamInfo +" with scoure " + firstTeamInfo.totalScore + ")");
	    	  }
	      }
	      
	     print("East Side team First Round winners:");
	     printWinner(eastSide);
	     print("");
	     print("");
	     print("West Side team First Round winners:");
	     printWinner(westSide);
	     
	}
	private static void printWinner(TreeMap tm){
	      Set<Entry<Integer, String>> set = tm.entrySet();
	      // Get an iterator
	      Iterator i = set.iterator();
	      int j=0;
	      // Display elements
	      while(i.hasNext()) {
	    	  j++;
	        Map.Entry<Integer, String> me = (Map.Entry<Integer, String>)i.next();
	        print(me.getValue());
	      }	
	}
	
	private static void print(String x){
		System.out.println(x);
	}
	private static void calculateWeight() {
		WeightInfo weightInfo = null;
		TreeMap<Double, String> forPrint = new TreeMap<Double, String>();
		Iterator itor = datahs.entrySet().iterator();
		while (itor.hasNext()){
			 Map.Entry<String, TeamInfo> entry = (Entry<String, TeamInfo>) itor.next();
			 String teamName = entry.getKey();
			 TeamInfo ti = entry.getValue();
			 double defEff = ti.defEff;
			 weightInfo = weightHs.get("DefEff");
			 if (weightInfo.lowHigh){
				ti.defEffWeight = ((1 - ((defEff - ldefEff)/(hdefEff - ldefEff)) )*weightInfo.weight/100);
			 }else{
				 ti.defEffWeight = ((defEff - ldefEff)/(hdefEff - ldefEff))*weightInfo.weight/100;
			 }
			 double offEff = ti.offEff;
			 weightInfo = weightHs.get("OffEff");
			 if (weightInfo.lowHigh){
				ti.offEffWeight = ((1 - ((offEff - loffEff)/(hoffEff - loffEff)) )*weightInfo.weight/100);
			 }else{
				 ti.offEffWeight = ((offEff - loffEff)/(hoffEff - loffEff))*weightInfo.weight/100;
			 }
			 
			 double freeThrow = ti.freeThrow;
			 weightInfo = weightHs.get("FreeThrow");
			 if (weightInfo.lowHigh){
				ti.freeThrowWeight = ((1 - ((freeThrow - lfreeThrow)/(hfreeThrow - lfreeThrow)) )*weightInfo.weight/100);
			 }else{
				 ti.freeThrowWeight = ((freeThrow - lfreeThrow)/(hfreeThrow - lfreeThrow))*weightInfo.weight/100;
			 }
			 int greatCoach = ti.greatCoach;
			 weightInfo = weightHs.get("GreatCoach");
			 if (weightInfo.lowHigh){
				ti.greatCoachWeight = ((1 - greatCoach)*weightInfo.weight/100);
			 }else{
				 ti.greatCoachWeight = (greatCoach)*weightInfo.weight/100;
			 }
			 int starPlayer = ti.starPlayer;
			 weightInfo = weightHs.get("StarPlayer");
			 if (weightInfo.lowHigh){
				ti.starPlayerWeight = ((1- starPlayer)*weightInfo.weight/100);
			 }else{
				 ti.starPlayerWeight = (starPlayer)*weightInfo.weight/100;
			 }
			 double sos = ti.sos;
			 weightInfo = weightHs.get("Sos");
			 if (weightInfo.lowHigh){
				ti.sosWeight = ((1 - ((sos - lsos)/(hsos - lsos)) )*weightInfo.weight/100);
			 }else{
				 ti.sosWeight = ((sos - lsos)/(hsos - lsos))*weightInfo.weight/100;
			 }
			 double tournamentRank = (double)ti.tournamentRank;
			 weightInfo = weightHs.get("TournamentRank");
			 if (weightInfo.lowHigh){
				ti.tournamentRankWeight = ((1 - ((tournamentRank-1)/(htournamentRank - 1)) )*weightInfo.weight/100);
			 }else{
				 ti.tournamentRankWeight = ((tournamentRank-1)/(htournamentRank - 1))*weightInfo.weight/100;
			 }
			 double kenpomRank = (double)ti.kenpomRank;
			 weightInfo = weightHs.get("KenpomRank");
			 if (weightInfo.lowHigh){
				ti.kenpomRankWeight = ((1 - ((kenpomRank-1)/((double)hkenpomRank - 1)) )*(double)weightInfo.weight/100);
			 }else{
				 ti.kenpomRankWeight = (((double)kenpomRank -1)/((double)hkenpomRank - 1))*(double)weightInfo.weight/100;
			 }
			 teamRanks.put(ti.kenpomRank, teamName);
			 double turnoverMargin = ti.turnoverMargin;
			 weightInfo = weightHs.get("TurnoverMargin");
			 if (weightInfo.lowHigh){
				ti.turnoverMarginWeight = ((1 - ((turnoverMargin - lturnoverMargin)/(hturnoverMargin - lturnoverMargin)) )*weightInfo.weight/100);
			 }else{
				 ti.turnoverMarginWeight = ((turnoverMargin - lturnoverMargin)/(hturnoverMargin - lturnoverMargin))*weightInfo.weight/100;
			 }
			 double offReb	= ti.offReb;
			 weightInfo = weightHs.get("OffReb");
			 if (weightInfo.lowHigh){
				ti.offRebwWeight = ((1 - ((offReb - loffReb)/(hoffReb - loffReb)) )*weightInfo.weight/100);
			 }else{
				 ti.offRebwWeight = ((offReb - loffReb)/(hoffReb - loffReb))*weightInfo.weight/100;
			 }
			 ti.totalScore = ti.offRebwWeight + ti.defEffWeight + ti.freeThrowWeight + ti.greatCoachWeight + ti.starPlayerWeight + ti.kenpomRankWeight + ti.sosWeight + ti.tournamentRankWeight
					 + ti.turnoverMarginWeight + ti.offEffWeight;
			 ti.totalScore = Math.round(ti.totalScore*100000);
			 datahs.put(teamName, ti);
			 forPrint.put(ti.totalScore, " team Name: " + teamName + " total score: " + ti.totalScore);
			 //res += " team Name: " + teamName + " total score: " + ti.totalScore + "\n";
			 //System.out.println(ti.toString());
		}
		printWinner(forPrint);
		print("");
		//System.out.println(res);
	}

	static void getDataFile(){
		File f = new File("");
		try {
			String tempStr = f.getCanonicalPath();
			System.out.println("data folder is "+tempStr);
			File[] fList = new File(tempStr).listFiles();
			for (File tmpf : fList){
				//System.out.println(tmpf.getName());
				if (tmpf.isFile()){
					String str = tmpf.getName();
					if ((str.endsWith(".xlsx") || str.endsWith(".xls")) &&(str.toLowerCase().contains("data"))){
						dataFile=str;
						break;
					}
				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	public static void getInputData() throws IOException
	{

		String ret="";
		try
		{
			//       FileInputStream file = new FileInputStream(new File("howtodoinjava_demo.xlsx"));
			FileInputStream file = new FileInputStream(new File(dataFile));

			//Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			workbook.setMissingCellPolicy(Row.CREATE_NULL_AS_BLANK);
			workbook.setMissingCellPolicy(Row.RETURN_NULL_AND_BLANK);
			//Get first/desired sheet from the workbook
			int totsheets = workbook.getNumberOfSheets();
			Iterator<XSSFSheet> it =workbook.iterator();
			XSSFSheet catSheet =null;
			XSSFSheet dataSheet =null;
			while(it.hasNext()){
				XSSFSheet sheet = it.next();
				if (sheet.getSheetName().toLowerCase().contains("weight")){
					catSheet =sheet;
				}else if (sheet.getSheetName().toLowerCase().contains("data")){
					dataSheet = sheet;
				}
			}
			loadCategory(catSheet);
			loadTeamData(dataSheet);

			file.close();
		} 
		catch (Exception e) 
		{
			e.printStackTrace();
		}finally {

		}

	}		
	private static void loadTeamData(XSSFSheet sheet) {
		Iterator<Row> rowIterator = sheet.iterator();
		boolean headerFound = false;
		boolean dataFound =false;
		boolean orderFound =false;
		while (rowIterator.hasNext()) 
		{
			Row row = rowIterator.next();
			//System.out.println("first cell>>>" + row.getFirstCellNum() + " last cell" + row.getLastCellNum());
			//fill in the empty cell with empty space instead of null
			if (row.getFirstCellNum()>0){
				int firstcellnm = row.getFirstCellNum();
				int lastcellnm = row.getLastCellNum();
				for (int r=0;r<2*firstcellnm -lastcellnm -1;r++){
					row.createCell(r).setCellValue("");
				}

			}
			if (!headerFound){
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					//based on the font color to fetch the page id (only the red one will be picked)
					//if (cell.getCellStyle().getFillForegroundColor() != IndexedColors.RED.getIndex()) continue;
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
					case Cell.CELL_TYPE_STRING:
						String cellValue = cell.getStringCellValue();
						if (cellValue.trim().equalsIgnoreCase("team")){
							//System.out.println("team column indx>" + cell.getColumnIndex());
							weightListHs.clear();
							headerFound = true;
						}else{
							cellValue = convertToTitleCase(cellValue).replace(" ", "");
							//System.out.println(cellValue + " column indx>" + cell.getColumnIndex());
							weightListHs.put(cell.getColumnIndex(), cellValue);
						}
						break;

					}
				}
			}else{
				Iterator<Cell> cellIterator = row.cellIterator();
				String cellValue="";
				TeamInfo teamInfo=null;
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					//based on the font color to fetch the page id (only the red one will be picked)
					//if (cell.getCellStyle().getFillForegroundColor() != IndexedColors.RED.getIndex()) continue;
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
					case Cell.CELL_TYPE_STRING:
						 cellValue = cell.getStringCellValue();
						 teamInfo = datahs.get(cellValue);
						 if (teamInfo==null){
							 teamInfo = new TeamInfo();
							 teamInfo.teamName=cellValue;
						 }
						 datahs.put(cellValue, teamInfo);
					break;
					case Cell.CELL_TYPE_NUMERIC:
						cell.setCellType(Cell.CELL_TYPE_STRING);
						 String cellValuestr = cell.getStringCellValue();
						 String catKey2 = weightListHs.get(cell.getColumnIndex());
						 switch (catKey2)
						 {
							 case "KenpomRank":
								 teamInfo.kenpomRank = Integer.parseInt(cellValuestr);
								 if(hkenpomRank<teamInfo.kenpomRank) hkenpomRank=teamInfo.kenpomRank;
								 break;
							 case "DefEff":
								 teamInfo.defEff = Double.parseDouble(cellValuestr);
								 if(hdefEff<teamInfo.defEff) hdefEff=teamInfo.defEff;
								 if(ldefEff>teamInfo.defEff) ldefEff=teamInfo.defEff;
								 break;								 
							 case "OffEff":
								 teamInfo.offEff = Double.parseDouble(cellValuestr);
								 if(hoffEff<teamInfo.offEff) hoffEff=teamInfo.offEff;
								 if(loffEff>teamInfo.offEff) loffEff=teamInfo.offEff;								 
								 break;								 
							 case "Sos":
								 teamInfo.sos = Double.parseDouble(cellValuestr);
								 if(hsos<teamInfo.sos) hsos=teamInfo.sos;
								 if(lsos>teamInfo.sos) lsos=teamInfo.sos;
								 break;								 
							 case "FreeThrow":
								 teamInfo.freeThrow = Double.parseDouble(cellValuestr);
								 if(hfreeThrow<teamInfo.freeThrow) hfreeThrow=teamInfo.freeThrow;
								 if(lfreeThrow>teamInfo.freeThrow) lfreeThrow=teamInfo.freeThrow;
								 break;								 
							 case "OffReb":
								 teamInfo.offReb = Double.parseDouble(cellValuestr);
								 if(hoffReb<teamInfo.offReb) hoffReb=teamInfo.offReb;
								 if(loffReb>teamInfo.offReb) loffReb=teamInfo.offReb;
								 break;								 
							 case "TurnoverMargin":
								 teamInfo.turnoverMargin = Double.parseDouble(cellValuestr);
								 if(hturnoverMargin<teamInfo.turnoverMargin) hturnoverMargin=teamInfo.turnoverMargin;
								 if(lturnoverMargin>teamInfo.turnoverMargin) lturnoverMargin=teamInfo.turnoverMargin;
								 break;								 
							 case "StarPlayer":
								 teamInfo.starPlayer = Integer.parseInt(cellValuestr);
								 break;								 
							 case "GreatCoach":
								 teamInfo.greatCoach = Integer.parseInt(cellValuestr);
								 break;								 
							 case "TournamentRank":
								 teamInfo.tournamentRank = Integer.parseInt(cellValuestr);
								 if(htournamentRank<teamInfo.tournamentRank) htournamentRank=teamInfo.tournamentRank;
								 break;								 

						 }
			
						 datahs.put(cellValue, teamInfo);
						break;					
					}
				}
				
			}
		}	
	}

	private static void loadCategory(XSSFSheet sheet) {
		Iterator<Row> rowIterator = sheet.iterator();
		boolean headerFound = false;
		boolean dataFound =false;
		boolean orderFound =false;
		while (rowIterator.hasNext()) 
		{
			Row row = rowIterator.next();
			//System.out.println("first cell>>>" + row.getFirstCellNum() + " last cell" + row.getLastCellNum());
			//fill in the empty cell with empty space instead of null
			if (row.getFirstCellNum()>0){
				int firstcellnm = row.getFirstCellNum();
				int lastcellnm = row.getLastCellNum();
				for (int r=0;r<2*firstcellnm -lastcellnm -1;r++){
					row.createCell(r).setCellValue("");
				}

			}
			if (!headerFound){
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					//based on the font color to fetch the page id (only the red one will be picked)
					//if (cell.getCellStyle().getFillForegroundColor() != IndexedColors.RED.getIndex()) continue;
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
					case Cell.CELL_TYPE_STRING:
						String cellValue = cell.getStringCellValue();
						if (cellValue.trim().equalsIgnoreCase("cat")){
							//System.out.println("cat column indx>" + cell.getColumnIndex());
							headerFound = true;
						}else{
							cellValue = convertToTitleCase(cellValue).replace(" ", "");
							//System.out.println(cellValue + " column indx>" + cell.getColumnIndex());
							weightListHs.put(cell.getColumnIndex(), cellValue);
						}
						break;

					}
				}
			}else{
				Iterator<Cell> cellIterator = row.cellIterator();
				String cellValue="";
				while (cellIterator.hasNext()) 
				{
					Cell cell = cellIterator.next();
					//based on the font color to fetch the page id (only the red one will be picked)
					//if (cell.getCellStyle().getFillForegroundColor() != IndexedColors.RED.getIndex()) continue;
					//Check the cell type and format accordingly
					switch (cell.getCellType()) 
					{
					case Cell.CELL_TYPE_STRING:
						 cellValue = cell.getStringCellValue();
						 String catKey = weightListHs.get(cell.getColumnIndex());
						 if (catKey!=null){
							 WeightInfo weightInfo = weightHs.get(catKey);
							 if (weightInfo==null){
								 weightInfo = new WeightInfo();
								 weightInfo.weightName=catKey;
							 }
							 weightInfo.lowHigh = (cellValue.trim().equalsIgnoreCase("low-high")?true:false);
							 weightHs.put(catKey, weightInfo);
						 }
							
					break;
					case Cell.CELL_TYPE_NUMERIC:
						cell.setCellType(Cell.CELL_TYPE_STRING);
						 cellValue = cell.getStringCellValue();
						 String catKey2 = weightListHs.get(cell.getColumnIndex());
						 if (catKey2!=null){
							 WeightInfo weightInfo = weightHs.get(catKey2);
							 if (weightInfo==null){
								 weightInfo = new WeightInfo();
								 weightInfo.weightName=catKey2;								 
							 }
							 weightInfo.weight = Integer.parseInt(cellValue);
							 weightHs.put(catKey2, weightInfo);
						 }
						break;					
					}
				}
				
			}
		}
	}

	static String convertToTitleCase(String s){
	     String[] words = s.toLowerCase().split(" ");

	     //next step is to do the capitalizing for each word
	     //so use a loop to itarate through the array
	     for(int i = 0; i< words.length; i++){
	        //we will save the capitalized word in the same place again
	        //first, geht the character on first position 
	        //(words[i].charAt(0))
	        //next, convert it to upercase (Character.toUppercase())
	        //then add the rest of the word (words[i].substring(1))
	        //and store the output back in the array (words[i] = ...)
	        words[i] = Character.toUpperCase(words[i].charAt(0)) + 
	                  words[i].substring(1);
	     }

	    //now we have to make a string out of the array, for that we have to 
	    // seprate the words with a space again
	    //you can do this in the same loop, when you are capitalizing the 
	    // words!
	    String out = "";
	    for(int i = 0; i<words.length; i++){
	       //append each word to out 
	       //and append a space after each word
	       out += words[i] + " ";
	    }	
	    return out.trim();
	}
	
	static class TeamInfo{
		String teamName;
		int kenpomRank;
		double kenpomRankWeight;
		int tournamentRank;
		double tournamentRankWeight;
		double defEff;
		double defEffWeight;
		double offEff;
		double offEffWeight;
		double sos;
		double sosWeight;
		double freeThrow;
		double freeThrowWeight;
		double offReb;
		double offRebwWeight;
		double turnoverMargin;
		double turnoverMarginWeight;
		int starPlayer;
		double starPlayerWeight;
		int greatCoach;
		double greatCoachWeight;
		double totalScore;
		@Override
		public String toString() {
			return "TeamInfo [teamName=" + teamName + ", kenpomRank="
					+ kenpomRank + ", kenpomRankWeight=" + kenpomRankWeight
					+ ", tournamentRank=" + tournamentRank
					+ ", tournamentRankWeight=" + tournamentRankWeight
					+ ", defEff=" + defEff + ", defEffWeight=" + defEffWeight
					+ ", offEff=" + offEff + ", offEffWeight=" + offEffWeight
					+ ", sos=" + sos + ", sosWeight=" + sosWeight
					+ ", freeThrow=" + freeThrow + ", freeThrowWeight="
					+ freeThrowWeight + ", offReb=" + offReb
					+ ", offRebwWeight=" + offRebwWeight + ", turnoverMargin="
					+ turnoverMargin + ", turnoverMarginWeight="
					+ turnoverMarginWeight + ", starPlayer=" + starPlayer
					+ ", starPlayerWeight=" + starPlayerWeight + ", greatCoach="
					+ greatCoach + ", greatCoachWeight=" + greatCoachWeight
					+ ", totalScore=" + totalScore + "]";
		}		
	}
	static class WeightInfo{
		String weightName;//rank, def eff or catergory name
		int weight;//in percentage, whole number for now
		boolean lowHigh;//lower number higher the rank
		
	}
}
