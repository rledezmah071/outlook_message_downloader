# outlook_message_downloader

## Tool description

Tool created for GCR team to download latest emails based on the number of days requested from specified Outlook Folder. After downloading the emails, the tool will find phone number and MTCNs present in the body and save to an excel file.

## Requirements
* Download Anaconda Python 3 (64-bit version for Windows) on https://www.anaconda.com/products/individual#Downloads.
* Follow the Anaconda installation steps.
* Unzip the "outlook_message_downlader.zip" files to your computer.

## How to use
1. Open, with a text editor, the file settings.yml located in etc/ folder and set the tool configurations
2. From Windows Home Menu, look for "Anaconda Prompt (Anaconda3) and open it"
  1. Input "cd " followed by the location where you unzipped the program code and type Enter. Example = "cd C:\Users\XXXXXX\Documents\outlook_message_downloader"
  2. Input "python omd.py" to run the program and type Enter.
  3. The tool will download the results into the folder selected in tool configurations. (You'll notice the code is finished when you see a line blinking again in the console)
 
