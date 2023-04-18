<#PSScriptInfo
.VERSION 2023.4.18.1144
.GUID 6671c61e-aac0-46d6-a6ee-5d588d246df9
.AUTHOR rosenqui
.COPYRIGHT
    Copyright (c) 2023 Eric Rosenquist, https://github.com/rosenqui
.LICENSEURI
    https://github.com/rosenqui/oeb-rate-analyzer/blob/main/LICENSE
.PROJECTURI
    https://github.com/rosenqui/oeb-rate-analyzer
#>

#Requires -Version 5

<#
.SYNOPSIS
Looks at your hourly electrical consumption and computes the cost
using time-of-use (TOU), tiered, and ultra-low-overnight (ULO) rates
currently in effect.

.DESCRIPTION 
Analyze your electrical consumption based on Ontario Energy Board
electricity rates (https://www.oeb.ca/consumer-information-and-protection/electricity-rates).

Note that this utility only computes the cost of the "electricity charge"
portion of your electrical bill. It does not factor in the cost of the
"regulatory charges" or "delivery" portion of your bill. Those portions
are typically independent of your rate plan and vary based on your total
killowatt-hour usage.

Note also that this utility aggregates usage by month of the year,
whereas your bill most likely spans 30 or 31 days starting mid-month.
This would not have any impact on the time-of-use calculations, but
might have a small impact on the totals for tiered usage, especially
in billing periods that span the winter/summer cutover.

GETTING YOUR DATA
-----------------
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

.EXAMPLE
Analyze-ElectricalRate.ps1 -UsageCSV 2022-hourly.csv | Format-Table -AutoSize *

Computes the cost of each rate plan based on your 2022 usage data, aggregating
it by month and showing you which plan is the best for each month of the year.

The output contains the following columns:

 Month - the month of the year (1-12)
 IsWinter - does this month use winter or summer tiered rates?
 kWh - total kWh for the month
 Tiered - total cost (in dollars and cents) under the tiered rate plan
 Tier1kWh - number of kWh consumed at the tier1 rate
 Tier2kWh - number of kWh consumed at the tier2 rate
 TOU - total cost (in dollars and cents) under the time-of-use rate plan
 TOUkWhOffPeak - number of kWh consumed during off-peak hours
 TOUkWhMidPeak - number of kWh consumed during mid-peak hours
 TOUkWhPeak - number of kWh consumed during peak hours
 ULO - total cost (in dollars and cents) under the ultra-low-overnight rate plan
 ULOkWh - number of kWh consumed during ultra-low-overnight hours
 ULOkWhOffPeak - number of kWh consumed during ULO off-peak hours
 ULOkWhMidPeak - number of kWh consumed during ULO mid-peak hours
 ULOkWhPeak - number of kWh consumed during ULO peak hours
 Best - which of the three rate plans is the cheapest for the month

.EXAMPLE
Analyze-ElectricalRate.ps1 -UsageCSV 2022-hourly.csv -RawValues | Out-GridView

Augments the hour-by-hour data with per-tier values and then displays the
results in a GridView window. Use the -RawValues option to check that each
hourly usage bucket is being categorized correctly in terms of weekend / weekday /
holiday and the time-of-use rate plan.
#>

[CmdletBinding()]
[OutputType([pscustomobject])]
Param (
    # A CSV file giving hourly kWh usage in 3 columns:
    #  1) the date (ex. 2022-05-24 or 05/24/2022)
    #  2) the time of day at the start of the usage period. This should be an hour without any minutes, ex. 1:00 PM or 13:00
    #  3) the kWh usage for that hour
    #
    # The first two columns are combined and then parsed by the .NET DateTime::Parse method
    # (see https://learn.microsoft.com/en-us/dotnet/api/system.datetime.parse?view=netframework-4.8.1),
    # so they must combine to make a date+time format that .NET can parse.
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $UsageCSV,

    # If specified, show the computed values for each hour and don't roll
    # them up into monthly totals. This skips the analysis step, but lets you
    # double-check that the usage is being categorized correctly.
    [Parameter(Mandatory = $false)]
    [switch] $RawValues
)

Begin {
    $HOLIDAYS = @{   # Days of the year 
        2021 = @(1, 46, 92, 144, 182, 214, 249, 284, 359, 360);
        2022 = @(3, 52, 105, 143, 182, 213, 248, 283, 359, 360);
        2023 = @(2, 51, 97, 142, 184, 219, 247, 282, 359, 360);
    };

    $MID_PEAK = 'MidPeak';
    $OFF_PEAK = 'OffPeak';
    $ON_PEAK = 'Peak';
    $ULO = 'Ulo';
    $ULO_PEAK = 'UloPeak';

    $MidPeakRate = 0.102;
    $OffPeakRate = 0.074;
    $OnPeakRate = 0.151;

    $Tier1Rate = 0.087;
    $Tier2Rate = 0.103;
    $Tier2SummerThreshold = 600;
    $Tier2WinterThreshold = 1000;

    $UloRate = 0.024;
    $UloOffPeakRate = $OffPeakRate;
    $UloMidPeakRate = $MidPeakRate;
    $UloOnPeakRate = 0.24;

    function ComputeTieredRate($obj) {
        if ($obj.IsWinter) {
            $threshold = $Tier2WinterThreshold;
        } else {
            $threshold = $Tier2SummerThreshold;
        }

        if ($obj.kWh -le $threshold) {
            $t1 = $obj.kWh * $Tier1Rate;
            $t2 = 0;

            $obj.Tier1kWh = $obj.kWh;
            $obj.Tier2kWh = 0;
        } else {
            $tier2kWh = $obj.kWh - $threshold;

            $t1 = $threshold * $Tier1Rate;
            $t2 = $tier2kWh * $Tier2Rate;

            $obj.Tier1kWh = $threshold;
            $obj.Tier2kWh = $tier2kWh;
        }
        $obj.Tiered = $t1 + $t2;
    }

    function ComputeTouRate($obj) {
        $rate = 0;
        $kind = '';

        if ($obj.IsWeekend -or $obj.IsHoliday) {
            $rate = $OffPeakRate;
            $kind = $OFF_PEAK;
        } elseif ($obj.Hour -ge 19 -or $obj.Hour -lt 7) {
            $rate = $OffPeakRate;
            $kind = $OFF_PEAK;
        } elseif ($obj.Hour -ge 11 -and $obj.Hour -lt 17) {
            if ($obj.IsWinter) {
                $rate = $MidPeakRate;
                $kind = $MID_PEAK;
            } else {
                $rate = $OnPeakRate;
                $kind = $ON_PEAK;
            }
        } else {
            # We're in the 7-11AM and 5-7PM range
            if ($obj.IsWinter) {
                $rate = $OnPeakRate;
                $kind = $ON_PEAK;
            } else {
                $rate = $MidPeakRate;
                $kind = $MID_PEAK;
            }
        }
        $obj.Tou = $obj.kWh * $rate;
        $obj.TouRate = $rate;
        $obj.TouKind = $kind;
    }

    function ComputeUloRate($obj) {
        $rate = 0;
        $kind = '';

        if ($obj.Hour -ge 23 -or $obj.Hour -lt 7) {
            $rate = $UloRate;
            $kind = $ULO;
        } elseif ($obj.IsWeekend -or $obj.IsHoliday) {
            $rate = $UloOffPeakRate;
            $kind = $OFF_PEAK;
        } elseif ($obj.Hour -ge 16 -and $obj.Hour -lt 21) {
            $rate = $UloOnPeakRate;
            $kind = $ULO_PEAK;
        } else {
            $rate = $UloMidPeakRate;
            $kind = $MID_PEAK;
        }
        $obj.Ulo = $obj.kWh * $rate;
        $obj.UloRate = $rate;
        $obj.UloKind = $kind;
    }

    function IsHoliday([datetime]$dt) {
        $yearHolidays = $HOLIDAYS[$dt.Year];

        if ($null -eq $yearHolidays) {
            $false;
        } elseif ($yearHolidays -contains $dt.DayOfYear) {
            $true;
        } else {
            $false;
        }
    }
}

End {
    $dt = New-Object datetime;

    $data =
        Get-Content -Path $UsageCSV |
        ConvertFrom-CSV -Delimiter ',' -Header Date, Time, kWh |
        Foreach-Object {
            if ([datetime]::TryParse($_.Date + ' ' + $_.Time, [ref]$dt)) {
                [PSCustomObject]@{
                    Date = $dt;
                    kWh = [double]$_.kWh;

                    Month = $dt.Month;
                    DayOfWeek = $dt.DayOfWeek;
                    Hour = $dt.Hour;
                    IsHoliday = IsHoliday($dt);
                    IsWeekend = $dt.DayOfWeek -eq 'Saturday' -or $dt.DayOfWeek -eq 'Sunday';
                    IsWinter = $dt.Month -ge 11 -or $dt.Month -le 4;

                    Tou = 0.0;
                    TouRate = 0.0;
                    TouKind = '';

                    Ulo = 0.0;
                    UloRate = 0.0;
                    UloKind = '';
                }
            }
        } |
        Foreach-Object {
            ComputeTouRate($_);
            ComputeUloRate($_);
            $_;
        };

    if ($RawValues) {
        $data;
    } else {
        $data |
        Group-Object -Property Month |
        ForEach-Object {
            $obj = [PSCustomObject]@{
                Month = [int]$_.Name;
                IsWinter = [int]$_.Name -ge 11 -or [int]$_.Name -le 4;
                kWh = ($_.Group | Measure-Object -Sum kWh).Sum;
                Tiered = 0.0;
                Tier1kWh = 0.0;
                Tier2kWh = 0.0;
                TOU = ($_.Group | Measure-Object -Sum Tou).Sum;
                TOUkWhOffPeak = ($_.Group | Where-Object -Property TouKind -eq $OFF_PEAK | Measure-Object -Sum kWh).Sum;
                TOUkWhMidPeak = ($_.Group | Where-Object -Property TouKind -eq $MID_PEAK | Measure-Object -Sum kWh).Sum;
                TOUkWhPeak = ($_.Group | Where-Object -Property TouKind -eq $ON_PEAK | Measure-Object -Sum kWh).Sum;
                ULO = ($_.Group | Measure-Object -Sum Ulo).Sum;
                ULOkWh = ($_.Group | Where-Object -Property UloKind -eq $ULO | Measure-Object -Sum kWh).Sum;
                ULOkWhOffPeak = ($_.Group | Where-Object -Property UloKind -eq $OFF_PEAK | Measure-Object -Sum kWh).Sum;
                ULOkWhMidPeak = ($_.Group | Where-Object -Property UloKind -eq $MID_PEAK | Measure-Object -Sum kWh).Sum;
                ULOkWhPeak = ($_.Group | Where-Object -Property UloKind -eq $ULO_PEAK | Measure-Object -Sum kWh).Sum;
                Best = '';
            };

            ComputeTieredRate($obj);

            if ($obj.Tiered -lt $obj.Tou -and $obj.Tiered -lt $obj.Ulo) {
                $obj.Best = 'Tiered';
            } elseif ($obj.Tou -lt $obj.Tiered -and $obj.Tou -lt $obj.Ulo) {
                $obj.Best = 'TOU';
            } else {
                $obj.Best = 'ULO';
            }

            foreach ($field in ($obj | Get-Member -MemberType NoteProperty).Name) {
                if ($obj.$field -is [double]) {
                    $obj.$field = [System.Math]::Round($obj.$field, 2);
                }
            }

            $obj;
        }
    }
}
