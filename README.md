
# LinkedIn Message Automation  


## Table of Contents

- [Overview](#overview)
- [Getting Started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Installation](#installation)
- [License](#license)
- [Notes](#notes)


## Overview

The LinkedIn Message Automation project is designed to simplify the process of sending personalized messages to your LinkedIn connections. It includes features for automated messaging, tracking message statuses, and managing contact data through Excel files With the ability to cleanly stop the program at any given time using a specific keystroke.

## Getting Started  

  ### Prerequisites 
  Before you begin, ensure you have the following installed:  
- [Python](https://www.python.org/downloads/)
- [ChromeDriver](https://sites.google.com/chromium.org/driver/)  
- #### Set the following Environment variables in your OS:  
> > **LinkedInUserName** **:** yourEmailForLinkedin@gmail.com  
> > **LinkedinPassword** **:** ******  



## Installation

1. ###### Clone the repository:
    ```bash
    git clone https://github.com/OhadAms/Linkedin_Message_Automation.git

    
2. ###### Use the requirements.txt file added to the project to install all the needed packages:  
    ```bash
    pip install -r requirements.txt


3. ###### Or do it manualy:
    ```bash
    pip install pandas
    python -m pip install selenium 
    pip install deep_translator
    pip install tabulate 
    pip install openpyxl


## License
This project is licensed under the MIT License. See the [LICENSE](https://github.com/OhadAms/Linkedin_Message_Automation/blob/main/LICENSE) file for details.


## Notes    

1. #### An .xlsx file is used to store LinkedIn names along with corresponding hyperlinks to user profiles.

2. #### Create an .xlsx file.

3. #### Go to LinkedIn, and copy the LinkedIn username as follows:
   ![CopyLinkedinUserNameAsShownInImage](https://github.com/OhadAms/Linkedin_Message_Automation/blob/main/ReadMeImages/CopyLinkedinUserNameAsShownInImage.JPG)

   - ##### Make sure you can send that person a message on LinkedIn - (The person is in your connections or you have LinkedIn premium).

5. #### Paste the username into your working .xlsx file as follows:
   ![XlsxFileShouldLookLikeThis](https://github.com/OhadAms/Linkedin_Message_Automation/blob/main/ReadMeImages/XlsxFileShouldLookLikeThis.JPG)

6. #### In main.py, update the value of the 'originalXlsxFilePath' variable with the path to your original .xlsx file, as created in section '2'.

7. #### In main.py, update the value of the 'monitoringTableTargetXlsxFilePath' variable with the desired path where you want the Excel file containing summarized information to be saved.

8. #### In main.py, update the value of the 'maxMessageRound ' variable with the desired messages you want to send in each round of messaging.

9. #### In main.py, update the value of the 'nextMessagingRoundInSeconds' variable with the desired time for each round in seconds.
 
10. #### Run the program.

11. #### Wait until the program is finished running or stop the program manually using a long keystroke on 'ctrl+q' (note that the program will finish the current messaging round and stop only after that round is finishd, or if you try to stop it manually when its waiting for the next messaging round), to stop it manually use a long press on 'ctrl+q' until the program stops.

12. #### Go to your 'monitoringTableTargetXlsxFilePath' folder and open the .xlsx file created, the summarized information should be there as follows:
    ![summarizedXlsxFileShouldLookLikeThis](https://github.com/OhadAms/Linkedin_Message_Automation/blob/main/ReadMeImages/summarizedXlsxFileShouldLookLikeThis.JPG)
   
13. #### If you want to add more Usernames to the file after you ran the program once, just add them to the end of the original .xlsx file and run the program again.

14. #### Happy Job Hunting!




