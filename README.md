## NEC Tray Fill Calculation with Excel365

Calculating NEC tray fill with Excel365

## Required Data for Calculation

Note, all data are maintained as Excel table. There are columns for user to enter the data and the other columns contain formula of intermediate calculation.

*A. Data about the tray specification*

![Tray specification](src/images/TraySpecData.png?raw=true "Tray specification")

*B. Data about the cable specification*

![Cable specification](src/images/CableSpecData.png?raw=true "Cable specification")

*C. Data about trays*

![Cable trays](src/images/TrayData.png?raw=true "Cable trays")

Note, the result of the NEC tray fill calculation are shown in these columns as shown above:

- FillTotal
- RuleA
- RuleB
- RuleC

*D. Data about cables in trays*

![Cable in trays](src/images/CableData.png?raw=true "Cable in trays")

## Custom LAMBDA functions

The file [LAMBDA.txt](src/LAMBDA.txt) contains the code of custom LAMBDA functions:

- fxRuleA
- fxRuleB
- fxNecTable
- fxFillA
- fxFillB

## Abbreviation in column names and function arguments

- 1C_M2: single conductor size GTE1/0 and LTE4/0
- 1C_M3: single conductor size GTE250 and LTE900
- LG: large conductor GTE1000
- LV: insulation voltage LT2000
- MC: multi-conductor
- SA: sum of cable area
- SD: sum of cable diameter
- SIG: non-power, signal
- SM: small conductor size LT4/0
- NecTable: lookup values of NEC tables
- Rule: refer to NEC code section and sub-sections of tray fill calculation

## Miscellaneous

*A. Configurable values for dropdown validation*

![Configuration](src/images/Config.png?raw=true "Validation values")

*B. NEC lookup tables*

![NEC Tables](src/images/NECTable.png?raw=true "NEC Tables")

*C. NEC rule summary*

![NEC Rules](src/images/NECRule.png?raw=true "NEC Rules")

## Functional Programming Paradigm

This calculation was originally developed using C#. The C# source code is available at the ElectricalFunc/src/NecFillLib Github repository. The challenge is to re-implement this calculation in Excel365 without using VBA.

Logic in Excel can be implemented as Excel formula or as VBA macro. Excel365 has many built-in formula functions that perform logic similar to the procedural statements like:

- IFS() / SWITCH() for If-ElseIf-Else
- LAMBDA recursive call for looping
- LAMBDA / LET for custom function 
- Many functions for working with array of data
