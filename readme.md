# Excel Distribution Models VBA Script

This repository contains an Excel VBA script that provides various distribution models for forecasting purposes. The script includes functions to distribute a target value across a specified number of months using different mathematical models.

## Distribution Models

The script currently supports the following distribution models:

1. **Linear Distribution**: Equally distributes the target value across all months.
2. **Logarithmic Distribution**: Distributes the target value such that the values start low and increase gradually.
3. **Exponential Distribution**: Distributes the target value such that the values start low and increase rapidly.
4. **Normal Distribution**: Distributes the target value according to a normal distribution curve, peaking in the middle.
5. **Quadratic Distribution**: Distributes the target value such that the values start low and increase sharply towards the end, creating a "ski jump" shape.

## Installation

To use the VBA script, follow these steps:

1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module by going to `Insert` > `Module`.
4. Copy and paste the VBA code from this repository into the module.

## Usage

### Excel examples

Logarithmic distribution over 7 steps, show position 1
=DistributeGoal("logarithmic", 7, 1, $H$1)

quadratic distribution over 9 steps, show position 8
=DistributeGoal("quadratic", 9, 4, $H$1)

```vba
DistributeGoal(distributionType As String, totalMonths As Integer, currentPosition As Integer, target As Double) As Double
