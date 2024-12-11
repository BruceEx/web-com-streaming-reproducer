## Instructions:

### First create a file with a streaming formula using the OJS add-in
1. Use node version 18.18.0 (`nvm install 18.18.0 && nvm use 18.18.0`)
2. Install node modules (`npm install`)
3. Build app (`npm run build`)
4. Run node server by opening a terminal in project root and running `npm run dev-server`
5. Then run addin in excel by opening another terminal in project root: `npm run start:desktop`
6. PC Excel will open automatically
7. Type in Cell A1 `=TESTRTD(1,2)`, it would be streaming an incrementing value (ex. `OJS: 6`)
8. Save this file under `streaming-ojs`
9. Close the file

### Next we clear the OJS add-in from PC Excel
1. Go to the terminal running the node server and kill it
2. [Clear PC Excel Web add-in cache](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache#manually-clear-the-cache-in-excel-word-and-powerpoint) by deleting this folder %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\

### Next we install the COM add-in
1. Navigate to `the test-rtd-server` folder
2. Install the COM add-in by running `Install.bat` as admin
3. Open a new Excel file
4. Type in Cell A1 `=TESTRTD(1,2)`, check that it is streaming this value ex. `COM: #` (ex: `COM: 6`)
    - note that if the result it `#VALUE`:
      - make sure that you ran `Install.bat` as admin
      - close Excel and try opening it again
5. Close without saving

### Last we reproduce the bug by opening `streaming-ojs`
1. Open the previously saved excel file `streaming-ojs`
2. Notice that A1 which contains `=TESTRTD(1,2)` is not streaming and is stuck at with the cached web value `OJS: #` 
3. In A2 type `=TESTRTD(1,2)` and notice that it is streaming `COM: #`
4. Save and close excel file
5. Open it once again and notice that A2 resumes streaming but A1 is still stuck
6. Lastly enter A1, and notice it starts streaming using the COM formula: `COM: #`

#### Notes:
- source code for sample COM add-in is inside `test-rtd-server` folder
- the OJS code was generated using [yeoman generator in MS docs](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/yeoman-generator-overview)
