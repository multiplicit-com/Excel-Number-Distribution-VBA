# Excel Distribution Models VBA Script

This repository contains an Excel VBA script that provides various distribution models for forecasting purposes. The script includes functions to distribute a target value across a specified number of months using different mathematical models.

## Distribution Models

The script currently supports the following distribution models:

1. **Linear Distribution**: Equally distributes the target value across all months.
2. **Logarithmic Distribution**: Distributes the target value such that the values start low and increase gradually.
3. **Exponential Distribution**: Distributes the target value such that the values start low and increase rapidly.
4. **Normal Distribution**: Distributes the target value according to a normal distribution curve, peaking in the middle.
5. **Quadratic Distribution**: Distributes the target value such that the values start low and increase sharply towards the end, creating a "ski jump" shape.

![distribution_models](https://github.com/multiplicit-com/Excel-Number-Distribution-VBA/assets/127529943/663ca91f-99a5-4768-a535-4b08f842f1a6)

## Installation

To use the VBA script, follow these steps:

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module by going to `Insert` > `Module`.
4. Copy and paste the VBA code from this repository into the module.

Make sure the excel workbook is saved in the '.xlsm' file format, or it won't support VBA macros.

## Usage
Once installed, the **DistributeGoal** function can be called like any excel formula.

### Excel examples

Logarithmic distribution over 7 steps, show position 1:

 **_=DistributeGoal("logarithmic", 7, 1, $A$8)_**


quadratic distribution over 9 steps, show position 8:

 **_=DistributeGoal("quadratic", 9, 4, $A10)_**


### Parameters
DistributeGoal(distributionType As String, totalMonths As Integer, currentPosition As Integer, target As Double) As Double

* distributionType: the type of distribution model to apply to the target number.
  The accepted distribution types are:
  * linear
  * logarithmic
  * exponential
  * normal
  * quadratic
    
* totalMonths: The total number of months to distribute the goal across.
* currentPosition: The current position in the month range.
* target: The goal value to be distributed.
