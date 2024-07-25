This project contains 3 main part of auto test system: front end(choose which auto test will be execute), backend(call the test script),and auto test scripts.

To run the project:
1. Install PyCharm and Python 3
2. Install Flask, Angular, and Selenium in the project terminal
3. Download the code from GitHub, the link has been shared through email
4. Change the script file path(Country.py, Duedatelist.py) to the file path in your computer in the 'app.py' (Around line 36 and line 46).
5. Right click 'app.py' file and click 'run' button to start the backend
6. Switch to the 'Terminal' in the bottom of Pycharm, using git command 'cd' to get into the 'AngularUI' folder, type 'ng serve' to start the frontend
7. Open the browser, go to localhost:4200, choose 'Country App' or 'Report' under the 'Patent' dorpdown box, click 'Submit' button to run the auto test
8. Check if the auto test is running. If selected 'Country App' to run, check if there is an UAT form generated in the project after auto test finished.