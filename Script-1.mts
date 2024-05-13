'Open the Flight GUI Application
systemutil.Run "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\OpenText\OpenText UFT One\Sample Applications\Flight GUI"

'Checking if the screenshot file exists in the folder, if it does, then delete the file
Set objFso = CreateObject("Scripting.FileSystemObject")

destination_bitmap_file_path = "D:\Software Quality Control\Assignment 02 - UFT\screenshots\before_fiiling_login_info_" & DataTable("Iteration", dtGlobalSheet) & ".png"
If (objFso.FileExists(destination_bitmap_file_path)) Then
	objFso.DeleteFile(destination_bitmap_file_path)
End If

'Capture screenshot of main application before filling the login info
WpfWindow("OpenText MyFlight Sample").CaptureBitmap destination_bitmap_file_path 

'Fill in login details
WpfWindow("OpenText MyFlight Sample").WpfEdit("agentName").Set DataTable("Username", dtGlobalSheet)
WpfWindow("OpenText MyFlight Sample").WpfEdit("password").SetSecure "65c993b47b8a58479797"

'Checking if the screenshot file exists in the folder, if it does, then delete the file
destination_bitmap_file_path = "D:\Software Quality Control\Assignment 02 - UFT\screenshots\after_fiiling_login_info_" & DataTable("Iteration", dtGlobalSheet) & ".png"
If (objFso.FileExists(destination_bitmap_file_path)) Then
	objFso.DeleteFile(destination_bitmap_file_path)
End If

'Capture screenshot of main application after filling the login info
WpfWindow("OpenText MyFlight Sample").CaptureBitmap destination_bitmap_file_path

'Submit the login details by clicking OK button
WpfWindow("OpenText MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_2107286080_;_script infofile_;_ZIP::ssf4.xml_;_

'Checkpoint 1
WpfWindow("OpenText MyFlight Sample").WpfObject("John Smith").Check CheckPoint("TextCheckpoint")

'Manually added object WpfTabStrip and toggled tab selection
WpfWindow("OpenText MyFlight Sample").WpfTabStrip("WpfTabStrip").Select("SEARCH ORDER")
WpfWindow("OpenText MyFlight Sample").WpfTabStrip("WpfTabStrip").Select("BOOK FLIGHT")

'Checking if the screenshot file exists in the folder, if it does, then delete the file
destination_bitmap_file_path = "D:\Software Quality Control\Assignment 02 - UFT\screenshots\before_filing_flight_info_" & DataTable("Iteration", dtGlobalSheet) & ".png"
If (objFso.FileExists(destination_bitmap_file_path)) Then
	objFso.DeleteFile(destination_bitmap_file_path)
End If

'Capture screenshot before filling the flight details
WpfWindow("OpenText MyFlight Sample").CaptureBitmap destination_bitmap_file_path

'Fill in the flight details from GlobalData sheet (i.e., From City, To City, Flight Class, Number of Tickets, Flight Date )
WpfWindow("OpenText MyFlight Sample").WpfComboBox("fromCity").Select DataTable("From", dtGlobalSheet) @@ hightlight id_;_-3687112_;_script infofile_;_ZIP::ssf7.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfComboBox("toCity").Select DataTable("To", dtGlobalSheet) @@ hightlight id_;_2085083432_;_script infofile_;_ZIP::ssf9.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfImage("WpfImage_2").Click 10,15
WpfWindow("OpenText MyFlight Sample").WpfCalendar("datePicker").SetDate DataTable("Date", dtGlobalSheet)
WpfWindow("OpenText MyFlight Sample").WpfComboBox("Class").Select DataTable("Class", dtGlobalSheet) @@ hightlight id_;_2107294048_;_script infofile_;_ZIP::ssf13.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfComboBox("numOfTickets").Select DataTable("TicketCount", dtGlobalSheet) @@ hightlight id_;_-44637248_;_script infofile_;_ZIP::ssf17.xml_;_

'Checking if the screenshot file exists in the folder, if it does, then delete the file
destination_bitmap_file_path = "D:\Software Quality Control\Assignment 02 - UFT\screenshots\after_filing_flight_info_" & DataTable("Iteration", dtGlobalSheet) & ".png"
If (objFso.FileExists(destination_bitmap_file_path)) Then
	objFso.DeleteFile(destination_bitmap_file_path)
End If

'Capture screenshot after filling the flight details
WpfWindow("OpenText MyFlight Sample").CaptureBitmap destination_bitmap_file_path

'Checkpoint 2
WpfWindow("OpenText MyFlight Sample").WpfButton("FIND FLIGHTS").Check CheckPoint("StandardCheckpoint") @@ hightlight id_;_-20861672_;_script infofile_;_ZIP::ssf40.xml_;_

'Submit the flight details form by clicking Find Flights button
WpfWindow("OpenText MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_-44639072_;_script infofile_;_ZIP::ssf18.xml_;_

'Checkpoint 3
WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").Check CheckPoint("BitmapCheckpoint")

'Select the flight from result table
IterationCount = DataTable("Iteration", dtGlobalSheet)
If (IterationCount < 4) Then
   WpfWindow("OpenText MyFlight Sample").WpfTable("flightsDataGrid").SelectRow(1)	
End If

'Checkpoint 4 @@ hightlight id_;_-23752264_;_script infofile_;_ZIP::ssf47.xml_;_
WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").Check CheckPoint("FailureCheckpoint")

'Submit the flight selection by clicking Select Flight button
If(WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").GetROProperty("enabled", true)) Then
	WpfWindow("OpenText MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_-44637728_;_script infofile_;_ZIP::ssf20.xml_;_
Else
	WpfWindow("OpenText MyFlight Sample").Close
	ExitActionIteration
End If

'Wait for 2 seconds to make sure that Confirmation Page is loaded and CaptureBitmap doesn't capture the previous screen synchronously
Wait(2)

'Checking if the screenshot file exists in the folder, if it does, then delete the file
destination_bitmap_file_path = "D:\Software Quality Control\Assignment 02 - UFT\screenshots\before_fiiling_passenger_name_" & DataTable("Iteration", dtGlobalSheet) & ".png"
If (objFso.FileExists(destination_bitmap_file_path)) Then
	objFso.DeleteFile(destination_bitmap_file_path)
End If

'Capture screenshot before filling the passenger name
WpfWindow("OpenText MyFlight Sample").CaptureBitmap destination_bitmap_file_path

'Fill in the passenger's name
WpfWindow("OpenText MyFlight Sample").WpfEdit("passengerName").Set DataTable("Passenger", dtGlobalSheet) @@ hightlight id_;_-44637008_;_script infofile_;_ZIP::ssf21.xml_;_

'Checking if the screenshot file exists in the folder, if it does, then delete the file
destination_bitmap_file_path = "D:\Software Quality Control\Assignment 02 - UFT\screenshots\after_fiiling_passenger_name_" & DataTable("Iteration", dtGlobalSheet) & ".png"
If (objFso.FileExists(destination_bitmap_file_path)) Then
	objFso.DeleteFile(destination_bitmap_file_path)
End If

'Capture screenshot after filling the passenger name
WpfWindow("OpenText MyFlight Sample").CaptureBitmap destination_bitmap_file_path

'Complete the booking by clicking Order button
WpfWindow("OpenText MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_2085062888_;_script infofile_;_ZIP::ssf22.xml_;_

'End the flow by closing the application
WpfWindow("OpenText MyFlight Sample").Close @@ hightlight id_;_2102146120_;_script infofile_;_ZIP::ssf39.xml_;_
