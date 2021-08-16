# OWV-VBA
VBA work for tracking payroll

This program is used in Microsoft Excel to streamline the Payroll operation.

This is a 3-step system from employee to site manager to payroll director.

In this instance, a QR code is generated and used for the employees to clock in and out of work. Those submissions go to a google forms doc with the necessary information needed to run and organzie the payroll in Excel.

Excel was the requested destination so a VBA script was created to make sure there was no code in any cells to avoid accidental mistakes or deletions of code. 

Step.1 - Employee scans QR and fill out shift information
Step.2 - Manager audits then copies correct data from google forms and pastes onto Excel sheet, clicks button
Step.3 - Payroll director refreshes pivot

We pick up the process mid-Step.2

The manager copies the data from google forms, columns needed are; 
    Date - The date of the shift worked
    Employee - The employee who worked it
    Account (if applicable) - Which account that shift was worked at
    Role - The role of the employee (Attendant, Lead, Manager)
    Time-In - Arrival Time
    Time-Out - Departure Time

This is where we enter the VBA process.

The manager opens Excel and ENABLES MACROS!*important*

![](images/OWV-VBA_pic1.png)

The manager will then paste the google forms columns in column A of the Excel sheet under the orange colored cell. (The orange colored cell always indicates the end of the last submission/paste). (red arrow) Keep an eye on the grey totals bar the orange arrow is pointing to. This will go away and re-surface at the bottom.

![](images/OWV-VBA_pic2.png)

After this the manager can run the totals by clicking the 'Click to fill totals' button on the top (blue arrow). This will auto generate all totals for the shifts worked. 

![](images/OWV-VBA_pic3.png)

Notice the grey totals bar moved down. The next time you want to add data to this sheet, start again at the cell directly under the latest A cell colored orange.(red arrow)
