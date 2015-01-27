# Excel Exploratory Data Analysis Tool
A C# tool to assist in 'feeling' and exploring related yet poorly structured Excel files at scale. Gain basic statistical information on the structure of the data sets and an initial set of column headers for classification.

Works through Excel so you don't have to. Because as nice as CSV files are, sometimes data and environments are far from perfect.

## Requirements

- This console application should only require the installation of [R.NET](https://rdotnet.codeplex.com/) and MS Office.
- You may need to add the Microsoft.Office.Interop.Excel reference. Version 15 (Office 2013) is recommended.

## Setup Tips

### Program.cs
- Line 26 sets the directory for the files that should be read, which is also where the final output files will be created.
- Line 29 sets some base assumptions. These should be updated as test are run against the data set.
- Line 108 and 112 sets location of two training lists. Insert as many terms as you like, but place them in column A. The term in A1 will be treated as the dominant term to germinate the learning method. If results are poor, variate that term.
- trainlist is treating as strongly assumed information. looselist is treated as a weak assumption.

### ExcelEDA.cs and TermSampling.cs
- There are a variety of thresholds set to 70% in both of these files (represented by .7). This should be treated only as a base assumption and experimented with against your own data sets.

## Files
### Program.cs
- Does what you think it does. Also sets some base assumptions.

### Library.cs
- Establishes some data type classes for storing and moving data around.

### ExcelLib.cs
- A variety of Excel interop functions to make life easier.

### ExcelEDA.cs
- Two step analysis to attempt to find true UsedRange, for when UsedRange returns oversized ranges. Also establishes start row and column should conversion of Interop Array Objects to Excel COM Addresses become necessary.

### TermSampling.cs
- Learning method to establish column header names with a high level of confidence. You can start with a single term and explore from there, or feed in many terms with both strong and weak assumption. The output can then be turned over to a classifying method.
