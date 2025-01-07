NOTE this app version is 2.1

# Bing Search Automation Application (EXE Version)
This is a standalone executable (EXE) Bing Search Automation tool that automates performing searches on Bing using a list of search terms provided in a `.txt` file. The program does not require any additional installations or dependencies, as it comes packaged as a self-contained EXE file.

## Features
- Automates searches on Bing using a list of search terms.
- Option to run the tool in headless mode (no GUI).
- Configurable search delay between queries.
- Saves logs of the search operations for debugging and tracking.
- Configurable timeout settings and retry logic for searches.
- Open-source and free to use, with modifications under the conditions in the LICENSE.txt.

## Requirements
- Microsoft Edge WebDriver (or another compatible WebDriver for your browser).
- A `.txt` file containing search terms (one per line).

## Installation
1. Download the EXE file for the Bing Search Automation tool.
2. Download Microsoft Edge WebDriver from https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/ (or use another WebDriver of your choice).
3. Place the WebDriver file in the same directory as the EXE file.
4. Create or prepare your search terms file (`search.txt`), with each search term on a new line.
   
## Usage
1. Double-click the EXE file to launch the application.
2. The interface will prompt you to select the `search.txt` file (which contains the search terms) and the WebDriver file path.
3. Optionally, you can enable headless mode to run the program without a visible browser window.
4. Click "Start" to begin the automated search process.

## License
This project is licensed under the GNU General Public License v3.0. See the LICENSE.txt for more details.

## Author
© 2024 Refter. All rights reserved.
