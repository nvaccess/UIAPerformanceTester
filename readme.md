# UIA Performance Tester

This work in progress project is to create a test framework plus assets which will test the over all performance of UI automation implementations, specifically concentrating on content-heavy examples such as documents in web browsers (E.g. Microsoft Edge).

This test framework is based upon Microsoft's Test Authoring and Execution Framework (TAEF).

## Required Dependencies
* windows 10
* Microsoft Visual Studio 2017 Community
* Microsoft Windows Driver Kit

## Running the Tests
* Open The included solution in Microsoft visual Studio 2017 community
* bild the solution
* Close any currently open assistive technology such as a Screen Reader
* Run (with no debugging) by pressing control+f5
 * Wait for the tests to finish.
	* For those without vision, listen to the beeps:
		* Each test starts with a low beep, and ends with a high beep.
		* Once you hear the last high beep with no more low beep after, the tests have completed.
* View the status of each test in the console window left open.
* Press any key to dismiss the console window.
* A standard WTT-log file is also created in the output directory for further machine processing of the test results if required.
 
 ## Current tests
 ### documentName
 This is a basic test which opens a document in Microsoft Edge, and then fetches the name of the document, timing the specific call to fetch the name.
This test is really just an example which ensures everything is working okay.

### horizontalNavbar 
This test opens a document in Edge containing a horizontal navbar made up of a list of 7 links.
 It times how long it takes to gather all the content by walking the document with UI Automation, similarly to how a screen reader might gather information to speak the navbar as a line.
  