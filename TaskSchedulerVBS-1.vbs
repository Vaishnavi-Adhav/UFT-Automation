'Create the UFT application object
Set uftApp = CreateObject("QuickTest.Application") 

'Launch UFT
uftApp.Launch 

'Maximize and Make UFT visible
uftApp.Visible = True 
uftApp.WindowState = "Maximized"

'Open the UFT test
uftApp.Open "C:\Users\Srush\OneDrive\Documents\UFT One\GUI_Flight_Test_01\GUI_Flight_Test_01", True  

'Set run settings for the test 
Set uftTest = uftApp.Test 

'Continue test even though error occurs 
uftTest.Settings.Run.OnError = "NextIteration"

Set uftResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object
uftResultsOpt.ResultsLocation = "C:\Users\Srush\OneDrive\Documents\UFT One\GUI_Flight_Test_01\GUI_Flight_Test_01\Res100" ' Set the results location

uftTest.Run uftResultsOpt ' Run the test

'Close the test and Quit UFT
uftTest.Close 
 
uftApp.quit 

'Release the resources 
Set uftTest = Nothing 
Set uftApp = Nothing
Set uftAutoExportResultsOpts = Nothing
Set uftAutoExportSettings = Nothing
