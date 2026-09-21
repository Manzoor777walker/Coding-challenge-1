# Coding-challenge-1
# Statistical measure using sample data set
# Which statistical measure would be choose [ Mean, Median, or Mode] — From the given data set and Why
Here,The dataset includes a major outlier on 04-01-2026 with a sales amount of 300,000, while all other recorded sales range between 4,500 and 26,000 so I choose median. Formula used: [ =MEDIAN(C2:C11)].
# Create an Excel formula to calculate this metric only for Region = "South", ignoring blank sales values in the given data set
                To calculate this metric only for Region = "South", ignoring blank sales values .formula used:[ =AVERAGEIF(D2:D11,"South",C2:C11)]
# If management still insists on using Average, how would you modify the Excel formula to reduce the impact of extreme values without removing rows? 
=TRIMMEAN(FILTER(C2:C11,(D2:D11="South")*(C2:C11<>"")),0.35)
here Why 0.35 works
For 6 data points, Excel calculates: $6 * 0.35 = 2.1$.Excel rounds this down to the nearest even number, 
which is 2.
It drops 1 highest value (300,000) and 1 lowest value (5,200), giving you the correct trimmed average of 23,250. 
# break down this formula
FILTER(C2:C11, (D2:D11="South") * (C2:C11<>""))

D2:D11="South": Checks which cells in column D equal "South". Returns TRUE or FALSE.

C2:C11<>"": Checks that the cell in column C is not blank/empty. Returns TRUE or FALSE.

* (Multiplication): Evaluates as a logical AND. A row is kept only if both conditions are TRUE.

FILTER(...): Extracts only the numeric values from C2:C11 that satisfy both conditions.
