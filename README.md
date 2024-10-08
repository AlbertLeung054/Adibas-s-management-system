Project Information
This project is about Adibas, a well-known multinational sport firm founded in 1949  that sells sport shoes, sport clothing and other sporting accessories.
Adibas is a multinational corporation that sells a large amount and variety of products every day. Adibas's system allows the manager to check and modify information related to purchase record, staff information, and customer information. The aim is to constantly regulate and improve a company's business performance through an integrated data storage system. In this project, different worksheets contain different data sets that simulate different situations and generate useful reports for organization


Source Data Sets
1.	Customer Information
2.	Purchase Record
3.	Staff Information


Reporting tasks
❖	Detailed information about customer — useful for building a close relationship with customer and better practice for marketing
❖	Calculate the profit for the firm — useful for the corporation to manage the cost and revenue
❖	Evaluate the performance of sales — based on the rating that provided by the customer to evaluate the sales’ performance
❖	Sales record — monitor each category of products’ selling monthly records for better inventory management
❖	Staff record-calculate the salary and bonus according to their performance.


1.0 VBA APPLICATION GUIDE

Step 1: Accessing internal system
i. To access Adibas’s internal system, the button (cmdClick) needs to be clicked

Step 2: Sign in. 
i.The welcome page will be shown after clicking the button. The manager will be able to enter the account name and password to sign in. (Username: Hello Password: Hello)
ii. Manager can choose the visibility of password by clicking “Show password”, otherwise password will be encrypted by “*”.    
iii. “Login successful!” will be displayed if the account name and password matches with =“hello” and  ”hello” respectively. 
iv.  If the information entered is incorrect or empty, the user will not be able to access the system and the following message (The username or password is invalid) will be displayed

Step 3: Accessing Adibas System Home Page
i. After successful login, the manager needs to select 1) file type 2) task. File type includes “Purchase Record”, “Customer information”, and “Staff Information”, while the task includes “Check”, “Modify”. There are 3 and 2 option buttons for each file type and task respectively. Then the manager will be able to browse the files and upload the excel file needed. Also, the manager can simply click the “Generate Report” button in the top right corner to proceed to the userform that generates the reports.

Step 4: Browse the required excel file
i. After clicking the “Browse” button, the system will display this page to let the user upload the required file.

Step 5: After clicking the “Continue” button, the relevant page will be shown based on the file type and action they chose. 
i. If the manager chooses “Customer Information” and “Check”, then the “Customer Information Data '' userform (FrmCustomerInformationCheck) will be displayed.
ii. If the manager chooses “Customer Information” and “Modify”, then the “Member Information Entry Form '' userform (FrmCustomerInformationModify) will be displayed.
iii. If the manager chooses “Purchase Record” and “Check”, then the “Purchase Record Data Check'' userform (FrmPurchaseRecordCheck) will be displayed.
iv. If the manager chooses “Purchase Record” and “Modify”, then the “Purchase Record Entry Form'' userform (FrmPurchaseRecordModify) will be displayed.
v. If the manager chooses “Staff Information” and “Check”, then the “Staff Information Check'' userform (FrmStaffInformationCheck) will be displayed.
vi. If the manager chooses “Staff Information” and “Modify”, then the “Staff Information Entry Form” userform (FrmStaffInformationModify) will be displayed.


Userforms:

Userform 1: FrmCustomerInformationCheck
i.The user needs to enter the member name and click “Check”. Information such as “Living Area”, “Age”, “Communication”, “Member Type”, “Email Address”, “Phone number” attached to the customer can be accessed. 
ii. If the user wants to enter another member name, the previously entered name should be cleared first. This can be done through the “Reset” button. 
iii. To return to the Home Page, “Back” button can be clicked.

Userform 2: FrmCustomerInformationModify
i. The manager can enter information such as “Member Name”, “Member ID”, “Age”, “Phone Number”, “Email Address” in a TextBox and choose the “Living Area” from ListBox(lstLiving). One of the following will be selected: “HK”, “KLN”, “NT” for “Living Area”. There are 2 and 3 option buttons for “Gender” and “Member Type” respectively. Option buttons for “Gender”: Male or Female, and for “Member Type”: Silver, Gold, or Premium. For the “Communication” method, there are 3 CheckBox choices such as “Phone call”, “Email”, and “Never”.
ii. Only numeric values can be entered in the txtAge. Otherwise the system will display the following message(Please enter members only!)
iii. There are some restrictions when choosing the “Communication” method. “Email” and “Never” can not be selected at the same time, and also the system will not allow “Phone call” and “Never” to be selected together.
iv. If some or all the entries are left blank, the following message will be displayed(Please fill out all the required fields!)
v. By clicking the “Save” button, data entries will be saved to the Customer Information workbook.
vi. To return to the Home Page, “Return” button can be clicked.

Userform 3: FrmPurchaseRecordCheck
i. The Manager needs to select the Start Date and the End Date to get information such as Total Sales and Average Sales (within the same month). ComboBox(Start Date): CbxDay1, CbxMonth1, CbxYear1. ComboBox(End Date): CbxDay2, CbxMonth2, CbxYear2. Since our source data focuses on January transactions, the manager has to select January for the Month. 
ii. To display Total Sales, loop through the sales data and sum the sales within the selected date range. Average Sales is calculated through Total Sales divided by the number of the selected days range.
iii. To display Total Sales and Average Sales, “View” button should be clicked. Average Sales would be calculated.
iv. By clicking the “Reset” button, all the data will be cleared and the ComboBox will be reset.
v. To return to the Home Page, “Back” button can be clicked.

Userform 4: FrmPurchaseRecordModify
i. The manager can enter information such as “Date”(format: YYYY/MM/DD), “Time”(format: HH:MM:SS), “Member Name”, “Member ID” in a TextBox. There are 4 Option buttons for “Member Type”: Non-member, Silver, Gold, Platinum. For “Product Name”(cbxproduct1-cbxproduct5), “Salesperson”(cbxsalesperson), and “Payment”(cbxpayment), items from ComboBox will be selected. To select “Quantity” for each product, scroll bars (Sbr1-Sbr5) will be used. 
ii. The user can choose up to 5 products from ComboBox.
iii. Depending on the payment type, relevant commission applies. 
iv. If the “Member Type” will be chosen as “Non-member”, “Member ID” and “Member Name” entries will be restricted. 
v. To calculate “Income After Charge”, price of each product will be multiplied by the chosen quantity, and further multiplied by (1-commision). 
vi. To display the total Income After Charge, all these values will be summed up. “Total Amount” is the sum of each product’s prices and its quantities before the relevant commission is applied. The formulas shown below: 
vii. After clicking the “Save” button, “Total Amount” and “Income after Charge” will be calculated and displayed. 
viii.  Data entries will be saved to the Purchase Record workbook.

Userform 5: FrmStaffInformationCheck
i.The user needs to enter the “Staff ID” and click “Check”. Information such as “Staff Name”, “Age”, “Gender”, “Post”, “Salary Per Hour”, “Working Hours” attached to the customer’s ID can be accessed. 
ii. “Total Basic Salary'' will be calculated based on “Salary Per Hour” and “Working Hours” from the excel file. Total Basic Salary(lblBSDisplay) = Salary Per Hour * Working Hours.
iii. To obtain the data of “Bonus”, “Sales amount”, and “Total salary”( after bonus) relevant with entered staff information, the manager has to input the “Purchase Record.xlsx” file. 
1)	Firstly, locate the column that matches the StaffID input using For Loop. Then sum up the relevant sales person’s sales amount based on the corresponding row. 
2)	The Amount of “Bonus” will be based on their sales amount. The higher the sales amount, the higher the bonus. 
3)	Total Salary is summation of Basic Salary and Bonus.
iv. If the user wants to clear the information on the form the “Reset” button should be used.
v. To return to the Home Page, “Back” button can be clicked.

Userform 6: FrmStaffInformationModify
i. The manager can enter information such as “Staff Name”, “Staff ID”, “Age”, “Salary per hour”, “Salary per hour” in a TextBox. There are 2 and 3 option buttons for “Gender” and “Post” respectively. Option buttons for “Gender”: Male or Female, and for “Post”: Manager, Full Time Staff, or Part Time Staff. 
ii. Only numeric values can be entered in the txtAge. Otherwise the system will display the following message(Please enter numbers only!)
iii. After clicking the “Save” button , if some or all of the entries are left blank, the following message will be 
displayed(Please provide all the required information.)
iv. If all the information has been entered correctly, data entries will be saved to the Staff Information workbook.

Userform 7: FrmGenerateReport
i. The Manager needs to upload PurchaseRecord.xlsx by clicking “Browse”, and can choose to generate: 1) Sales Record, 2) Staff Performance pivot tables.  The StaffInformation.xlsx file has to be uploaded to generate the Staff Information pivot table.
ii. After browsing the files, mpgBarChartSalesPerformance Multi Page Control will be enabled, and the user can select which table to generate: 1) Line Graph(Sales), 2) Bar Chart(Product Sales), 3) Bar Chart(Staff Performance). Refedit Control (RefEdit3) will be used to let the user select the range of cells.
