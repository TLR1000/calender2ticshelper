# calender2ticshelper

Code that will list relevant calendar entries to help filling out tics.
Funtion to run: GetEntriesForThisWeek

You may choose to run it as a macro or run it automatically.

As a macro:

To set up your code to execute as a macro in Outlook, follow these steps:

Open Outlook.

Press Alt + F11 to open the Visual Basic Editor.

In the Visual Basic Editor, insert a new module by selecting "Insert" from the menu bar, and then selecting "Module".

Copy and paste your VBA code into the new module.

Save the macro by selecting "File" from the menu bar, then selecting "Save".

In Outlook, go to "File" and select "Options".

In the Options window, select "Customize Ribbon" from the left-hand menu.

Click "New Group" to create a new group on the ribbon.

Select "Macros" from the "Choose commands from" dropdown menu.

Select the macro you just created from the list of available macros, and click "Add" to add it to the new group.

Click "OK" to save the changes and close the Options window.

The macro is now available in the new group on the ribbon. You can run the macro by selecting it from the ribbon.

Note: Macros in Outlook are typically triggered manually by the user, either by selecting them from the ribbon or by assigning them to a keyboard shortcut. It is also possible to trigger macros automatically based on certain events, but this requires additional programming and configuration.

Run Automatically:

You can schedule a piece of VBA code to run automatically at a specific time using the Windows Task Scheduler. Here's an overview of the process:

Open the Windows Task Scheduler by searching for "Task Scheduler" in the Start menu or by typing "taskschd.msc" in the Run dialog box (Windows Key + R).

In the Task Scheduler, click on "Create Basic Task" in the "Actions" pane on the right-hand side.

Follow the prompts to create a new basic task. Give the task a name and description, and specify the schedule for the task (e.g. daily, weekly, monthly, etc.). Set the start time for the task to the time when you want the VBA code to run.

When prompted to choose an action for the task, select "Start a program" and click "Next".

In the "Program/script" field, enter the full path to the program that will run the VBA code. This may be the path to Excel or another Office application that can execute VBA code. For example, the path to Excel may be "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE".

In the "Add arguments" field, enter the command-line arguments needed to run the VBA code. This will typically include the path to the Excel workbook that contains the VBA code and the name of the macro to run. For example, if the VBA code is in a workbook called "MyWorkbook.xlsm" and the macro is called "MyMacro", the command-line arguments might look like this: "C:\MyFolder\MyWorkbook.xlsm" /e MyMacro

Click "Finish" to create the task. You can then test the task by right-clicking on it in the Task Scheduler and selecting "Run".

Note that in order for this to work, you will need to make sure that Excel is installed on the computer where the task will run, and that the VBA code is stored in a workbook that can be opened by Excel. Additionally, you may need to configure the security settings in Excel to allow VBA code to run without user intervention.
