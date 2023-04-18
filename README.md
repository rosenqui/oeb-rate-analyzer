# Analyze-ElectricalRate.ps1

A PowerShell script to help you determine which electrical rate plan is the
most economical based on your past usage.

The script looks at your hourly electrical consumption and computes
costs for the three different rate plans based on current
[Ontario Energy Board electricity rates](https://www.oeb.ca/consumer-information-and-protection/electricity-rates).

Note that this utility only computes the cost of the "electricity charge"
portion of your electrical bill. It does not factor in the cost of the
"regulatory charges" or "delivery" portion of your bill. Those portions
are typically independent of your rate plan and vary based on your total
killowatt-hour usage and neighbourhood designation.

Note also that this utility aggregates usage by month of the year,
whereas your bill most likely spans 30 or 31 days starting mid-month.
This would not have any impact on the time-of-use calculations, but it
might have a small impact on the totals for tiered usage, especially
in billing periods that span the winter/summer cutover.

## GETTING YOUR DATA

You can download your hourly usage data from Hydro Ottawa by signing
into your account, then clicking on the "Billing" tab, and then at
the top of the screen under the "Usage" menu selecting "Download My Data".

Select "Hourly" and then click on "Change Date" and specify Jan 1 2022
as the "From" date and Dec 31 2022 as the "To" date. Click the "Submit"
button to set the date range, then click on the green Microsoft Excel
icon to export your data as an Excel file. You don't need to select a
full year's worth of data, but it's important to include full months,
otherwise the calculations for the tiered rate may end up being incorrect.

You can convert the downloaded Excel file to a CSV file for free at
https://cloudconvert.com/. The resulting file can be used directly by
this script.

## Sample Output

Month|IsWinter|kWh|Tiered|Tier1kWh|Tier2kWh|TOU|TOUkWhOffPeak|TOUkWhMidPeak|TOUkWhPeak|ULO|ULOkWh|ULOkWhOffPeak|ULOkWhMidPeak|ULOkWhPeak|Best
-|-|-|-|-|-|-|-|-|-|-|-|-|-|-|-
1|True|951.38|82.77|951.38|0|88.28|612.33|167.99|171.06|91.7|244.45|256.38|299.08|151.47|Tiered
2|True|873.31|75.98|873.31|0|80.83|564.91|153.97|154.43|85.22|220.01|227.72|282.95|142.63|Tiered
3|True|893.3|77.72|893.3|0|84.89|536.61|177.14|179.55|89.62|238.5|169.95|326.42|158.43|Tiered
4|True|894.23|77.8|894.23|0|82.83|576.03|160.01|158.19|88.58|239.66|191.42|307.93|155.22|Tiered
5|False|991.02|92.48|600|391.02|90.44|660.07|170.96|159.99|94.93|281.91|239.1|306.78|163.23|TOU
6|False|1211.92|115.23|600|611.92|110|812.85|212.39|186.68|123.28|350.04|243.37|373.73|244.78|TOU
7|False|1219.77|116.04|600|619.77|109.43|843.58|200.02|176.17|118.34|368.74|282.5|346.8|221.73|TOU
8|False|1294.93|123.78|600|694.93|116.82|885.75|214.53|194.65|130.07|423.79|235.67|362.74|272.73|TOU
9|False|1034.42|96.95|600|434.42|93.95|690.72|184.91|158.79|101.73|290.36|233.92|325.96|184.18|TOU
10|False|838.09|76.72|600|238.09|77.92|531.59|157.15|149.35|85.63|215.86|189.62|271.05|161.56|Tiered
11|True|890.78|77.5|890.78|0|84.79|532.94|177.16|180.68|92.78|226.4|172.87|314.57|176.94|Tiered
12|True|960.12|83.53|960.12|0|90.99|580.07|190.36|189.69|99.7|260.09|162.76|344.44|192.83|Tiered
