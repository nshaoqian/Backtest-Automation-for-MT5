# Backtest-Automation-for-MT5

To automate, simply run the python file "run-this-to-automate.py" and input required parameters.

Below is the breakdown:

----------------------------------------------------------------------------------------
0) Setup
- change directories terminal_folder, report_folder, opt_ini_folder, for_ini_folder 
- enter symbol (capital letter), year and number of training days

1) Write bat files for automations.
- automate-optimization.bat contains prompts to run optimization for each weeks
- automate-forecast.bat contains prompts to forecast profit for each weeks.

2) Write optimization ini files with different training periods.

3) Start automating optimization. 
- run automate-optimization.bat
- this process will take a long time as the number of optimization equals the number of weeks in the training period.
- once the optimization is done, the report folder will contain xml files containing data for forecasting.

4) Write forecasting ini files with different forecasting periods.
- read the xml files and extract optimized parameter values, then write it in each ini files

5) Start automating forecasting.
- run automate-optimization.bat
- once the forecasting is done, htm files with images about the results will be added to the report folder.

6) Saving results.
- extract weekly net profit from the htm files
- save the result into excel

----------------------------------------------------------------------------------------

remark: "XML structure.txt" depicts the hierachy of the data tree of the optimization report.
