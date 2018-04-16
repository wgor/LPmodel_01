This is a simple model of a sceduling LP optimization problem for electric energy prosumers that uses pv modules and battery storages. 
The model therefore solves a mixed integer minimization problem in order to compute the fewest energy costs for a prosumer over a 
given timeframe (96 steps), with given constraints and a chosen timeseries for market and energy feed-in prices. The user can choose 
from different demand and generation profiles and can configure a few other, mostly battery related, parameters for each prosumer.

Data input and data visualization needs to be done via the attached Excel file, whereas for the actual relaxation pythons pulp package 
is used. To start the solving process the user has to call the run.py file from console within the downloaded directory. 
Data between Excel and python is passed via pythons xlwings lib, while a discrete time simulation framework is provided by pythons 
lib mesa. For the exact package requirements please check out the attached requirements file.

Instructions
1. Open the Excel file within the file directory and put in some data
2. Leave the file open and start run.py from console within the file directory
3. Click plot buttons within Excelsheet to visualize simulation results
