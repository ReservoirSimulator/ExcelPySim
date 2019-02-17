# ExcelPySim
A reservoir simulation processing software based on the combination of Excel and Python. ![main banner](material/pythonExcel.png?raw=True "main banner")

## What
It's a reservoir simulation processing software based on Microsoft Excel, but is developed with Python language. That's, Microsoft Excel is the main work platform or main frame for reservoir simulation work, but the software is  developed with Python, not the embedded VBA. 
## Why
Reservoir or geological engineers are born data analyst！And Python is the popular language for data analysis, machine learning and AI, refer to [How popular is Python](https://www.quora.com/How-popular-is-python "Python is popular"). At the same time, Microsoft Excel is the software we use frequently when we work. Why not write a reservoir simulation processing software based on Excel+Python, which will enable us to learn new python skill while working in a familiar Excel environment？
## How
### *How to integrate Python into Excel?*  
Thanks [XlWings](https://github.com/ZoomerAnalytics/xlwings "XlWings"), it provides the solution for us because it makes it easy to call Python from Excel and vice versa. Please install [XlWings](https://github.com/ZoomerAnalytics/xlwings "XlWings") and read its documents at first before you try ExcelPySim. 
### *How it work?*   
A xlsm file is open to input data which is read by and converted into a deck file for XXSim reservoir simulation software. After finishing input, a button in the xlsm file could be clicked to call the XXSim simulator to run. Please visit [XXSim Github](https://github.com/ReservoirSimulator/XXSim "xxsim") or [XXSim Website](https://www.peclouds.com "xxsim website") for more.  
## Getting Started
It is started with a simple black oil case.
### Tutorial 1   
![tutor 1](material/tutor1.2019-02-14.gif?raw=True "tutor1 dynamic")
There are two files, 'tutor1.xlsm' and 'tutor1.py' which process the grid section of black oil case.  By tutorial 1, it is shown that the combination of [XlWings](https://github.com/ZoomerAnalytics/xlwings "XlWings") and Excel can provide a solution to a Excel-based reservoir simulation pre-processing tool. The use of xlwings is simple, just call your Python functions in your VBA macros. As follow is the python code to read and write Excel,
```python
        # read value from excel
        nx = sheet.range('B3').options(numbers=int).value
        # write value to excel
        sheet.range('A10').value = 'Grid Size: DX'
```

### Tutorial 2
![tutor 2](material/tutor2.gif?raw=True "tutor2 dynamic")
Reservoir engineers need to process production data frequently， so the reading and plotting of well production data is put forward to the tutorial 2. The production data format in tutor2.xlsm is just an example, so when you use it, it is need to adjust the Excel format and edit the code according to the actual situation. Two important libs for data analysis, numpy and pandas, have been imported in Tutorial 2 to process chunks of production data, which is much more convenient and effective than VBA. Just one line code to read production data:
```python
    df_prod = sheet.range((iRow,2),(iEndRow,7)).options(pd.DataFrame,index=False).value
```
Please visit [pandas tutorial](https://pandas.pydata.org/pandas-docs/stable/getting_started/tutorials.html "Pandas") to learn more.
### Tutorial 3
In this tutorial, PVT and SCAL data are readed and shown in graph. Also [Pandas](https://pandas.pydata.org "Pandas") is used to read and store table data.
## Joining us?
Reservoir simulation technology is important one of reservoir engineering methods and reservoir simulation tool should be one of tools at hand of a reservoir engineer. I hope it is helpful by the open source method to those that want to learn or use reservoir simulation technology and even python or machine learning. Anyone is welcome to participate.  
### About me
**Wang Tao, Co-Founder of XXSim reservoir simulation software**   
*Reservoir simulation engineer, C++ coder*  
Found more about Wang Tao check out these links:  
[LinkedIn](https://www.linkedin.com/in/tao-wang-xxsim/ "Linkedin") | [AAPG Blog](https://www.aapg.org/publications/blogs/learn/article/Articleid/42130/interview-with-tao-wang-pecloud "aapg") | [Twitter](https://twitter.com/wangtao74 "Twitter") | [Website](https://www.peclouds.com "peclouds")
