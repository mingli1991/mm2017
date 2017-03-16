This is a simple project to calculate the based on defined categories (10 at this time) with different weight.
The score are arbitary number, the high the better.

Basic design ideas:
1. read statistic data of each category for the teams and weight number for each category in a excelsheet.
2. based on the weight order (low number has higher weight or not) to calculate the percentage of the score for category:
	for example: def efficiency shorted as defEff. if the high/low of scores within all the team are 179/38, 
	weight for this category is 5 percent and the weight order is high-low (high score is better).
	For a team with score of 120, the contribution to winning the match for this category should be ((2.86% out of 5%)):
		((120 -40)/(180-40))*5/100 = (8/14)*0.05 = 0.2/7 ~ 0.0286

3. add the calculated the percentage for each category to get total score for each team (this total should be less than 1). 
4. print out the total score for each team for picking up the winner if any team meets. For readabilty, time the total with 100000.		