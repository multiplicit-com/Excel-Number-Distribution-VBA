# Excel Distribution Models VBA Script

This VBA macro script for Excel provides various distribution models for forecasting purposes. Simply provide a target number and specify the desired number of steps in the formula. The script allows you to apply different mathematical models to change the distribution.

All the numbers at each step will add up to the the 'goal' you provided. When expressed as a percentage the percentage of each 'step' will add up to 100%.

Note - there is also a verion for Google Sheets, which you can find here: https://github.com/multiplicit-com/Google-Sheets-Number-Distribution-AppScript-formula

## Distribution Models

The script currently supports the following distribution models:

1. **Linear Distribution**: Equally distributes the target value across all months.
2. **Logarithmic Distribution**: Distributes the target value such that the values start low and increase gradually.
3. **Exponential Distribution**: Distributes the target value such that the values start low and increase rapidly.
4. **Normal Distribution**: Distributes the target value according to a normal distribution curve, peaking in the middle.
5. **Quadratic Distribution**: Distributes the target value such that the values start low and increase sharply towards the end, creating a "ski jump" shape.

![distribution_models](https://github.com/multiplicit-com/Excel-Number-Distribution-VBA/assets/127529943/ba33b90a-df10-4d72-a0cb-845f72149f7b)


## Installation

To use the VBA script, follow these steps:

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module by going to `Insert` > `Module`.
4. Copy and paste the VBA code from this repository into the module.

Make sure the excel workbook is saved in the '.xlsm' file format, or it won't support VBA macros.

## Usage
Once installed, the **DistributeGoal** function can be called like any excel formula.

### Parameters
_VBA declaration:_ DistributeGoal(**DistributionType** As String, **TotalMonths** As Integer, **CurrentPosition** As Integer, **Target** As Double) As Double

* **DistributionType**: the type of distribution model to apply to the target number.
  The accepted distribution types are:
  * linear
  * logarithmic
  * exponential
  * normal
  * quadratic
    
* **TotalMonths**: The total number of months to distribute the goal across.
* **CurrentPosition**: The current position in the month range.
* **Target**: The goal value to be distributed.


### Excel examples

Logarithmic distribution over 7 steps, show position 1:

 **_=DistributeGoal("logarithmic", 7, 1, $A$8)_**


quadratic distribution over 9 steps, show position 8:

 **_=DistributeGoal("quadratic", 9, 8, $A10)_**

