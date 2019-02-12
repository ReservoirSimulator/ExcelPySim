# ExcelPySim
A reservoir simulation processing software based on the combination of Excel and Python. ![main banner](material/pythonExcel.png?raw=True "main banner")

## What
It's a reservoir simulation processing software based on Microsoft Excel, but is developed with Python language. That's, Microsoft Excel is the main work platform or main frame for reservoir simulation work, but the software is  developed with Python, not the embedded VBA. 
## Why
Reservoir or geological engineers are born data analyst！And Python is the popular language for data analysis, machine learning and AI, refer to [How popular is Python](https://www.quora.com/How-popular-is-python "Python is popular"). At the same time, Microsoft Excel is the software we use frequently when we work. Why not write a reservoir simulation processing software based on Excel+Python, which will enable us to learn new python skill while working in a familiar Excel environment？
## How
### *How to integrate Python into Excel?*  
Thanks [XlWings](https://github.com/ZoomerAnalytics/xlwings "XlWings"), it provides the solution for us because it makes it easy to call Python from Excel and vice versa. Please instll [XlWings](https://github.com/ZoomerAnalytics/xlwings "XlWings") and read its documents at first before you try ExcelPySim. 
### *How it work?*   
A xlsm file is open to input data which is read by and converted into a deck file for XXSim reservoir simulation software. After finishing input, a button in the xlsm file could be clicked to call the XXSim simulator to run. Please visit [XXSim Github](https://github.com/ReservoirSimulator/XXSim "xxsim") or [XXSim Website](https://www.peclouds.com "xxsim website") for more.  
## Getting Started
It is started with a simple black oil case.   
### Tutorial 1   
There are two files, 'tutor1.xlsm' and 'tutor1.py' which process the grid section of black oil case.  
## Joining us?
Reservoir simulation technology is important one of reservoir engineering methods and reservoir simulation tool should be one of tools at hand of a reservoir engineer. I hope it is helpful by open source method to those that want to learn or use reservoir simulation technology and even python or machine learning. Anyone is welcome to participate.
### About me
**Wang Tao, Co-Founder of XXSim reservoir simulation software**   
*Reservoir simulation engineer, C++ coder*  
Found more about Wang Tao check out these links:  
[LinkedIn](https://www.linkedin.com/in/tao-wang-xxsim/ "Linkedin") | [AAPG Blog](https://www.aapg.org/publications/blogs/learn/article/Articleid/42130/interview-with-tao-wang-pecloud "aapg") | [Twitter](https://twitter.com/wangtao74 "Twitter") | [Website](https://www.peclouds.com "peclouds")
