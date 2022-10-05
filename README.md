# Audit_Report_Gen
Audit Report generator

1. Put "WXX_Line_Mode_Sorted.xlsx" and "WXX_Line_Mode_Auto.xlsx" of this week and last week
   in the folder /Audit_Report_Gen/Sorted.

2. Edit the "Input Settings" part of Audit_Report_Gen.py.
   
   *If you need to compare several lines of two weeks, for example, W34_Main with W33_Main, W34_STR with W33_STR,
   use "order1 = ['Main_Resume', 'STR_Cold']
        order2 = ['Main_Cold', 'STR_Disable']".

    Note: output_order is for the sheet name of output file.

   *If you want to compare the results of different product lines, set "Enable_Productlines_Compare = True".

3. Execute the python script. You can find WXX_Audit_Report.xlsx in the folder /Audit_Report_Gen/Audit_Report.

##### DO NOT delete or move the folders "All", "Audit_Report", "Sorted" #####

##### DO NOT delete or put other contents before the sheet "Sheet" in "WXX_Line_Mode_Sorted.xlsx" and "WXX_Line_Mode_Auto.xlsx"  #####
