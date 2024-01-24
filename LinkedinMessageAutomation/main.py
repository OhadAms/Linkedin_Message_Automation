from deep_translator import GoogleTranslator
import linkedinFunctions as lf

# Our original and finalized .xlsx Data Frame target file paths.
originalXlsxFilePath = r"FolderLocation\testFile.xlsx"
monitoringTableTargetXlsxFilePath = r"FolderLocation\testFileUpdated.xlsx"

try:
    # Get the driver object from the point of after the initial login.
    driver = lf.linkedInLogin()

    # Create our lists of names and there correlated hyperLinks from the xlsx file.
    firstNames, lastNames, nameLinks, statuses, timeApproachedList = lf.createNameAndHyperLinkLists(originalXlsxFilePath)

    # Create the initial DataFrame.
    df = lf.createDataFrame(firstNames, lastNames, nameLinks, statuses, timeApproachedList)

    i = 1
    dailyMessagesSent = 0
    maxMessageRound = 6
    nextMessagingRoundInSeconds = 36000  # 10 hours
    template = "היי "
    messageBody = ",זה הטקסט שאני ארשום פה לכולם! רציתי לספר לך שזה עובד!\n"
    
    # Our exit key to stop the program while its working is 'ctrl+q' - you can change the 'exitKey' as you wish.
    exitKey = 'ctrl+q'
    while i <= len(df["Name"]) and not lf.checkExitProgram(exitKey):
        if df["Status"][i] != "Approached":
            # Checking if the original .xlsx file current entrance has been 'Approached', if so - we don't sent the message again and thus don't enter the if statement.
            if lf.checkIfEntranceInOriginalXlsx(originalXlsxFilePath, i, "Approached", "B"):
                # 1. Translate name and create full message
                name = GoogleTranslator(source='en', target='iw').translate(df["Name"][i])
                message = template + name + messageBody
                # 2. preform click on person link
                lf.openLink(driver, df["LinkedIn Link"][i])
                # 3. preform click on person message button
                lf.openLinkedInUserMessageBox(driver)
                # 4. preform click on person message area to get focus.
                lf.clickMessageArea(driver)
                # 5. enter text template into person input field.
                lf.findMessageParagraphAndEnterMessageTemplet(driver, message)
                # 6. send message.
                lf.sendMessage(driver)
                # 7. Add a daily MessageSent
                dailyMessagesSent += 1
                # 8. set 'Status' to "Approached".
                df["Status"][i] = "Approached"
                # Update original .xlsx file 'Status' to "Approached"
                lf.updateOriginalXlsxFile(originalXlsxFilePath, i, "Approached", "B")
                # 9. set 'Time Approached' to the current time that the template message was sent.
                df["Time Approached"][i] = lf.getDateAndTime()
                # Update original .xlsx file 'Time Approached' to the current time that the template message was sent.
                lf.updateOriginalXlsxFile(originalXlsxFilePath, i, df["Time Approached"][i], "C")
                # 10. Updating the monitoring Table Target .Xlsx File created with pandas.
                df.to_excel(monitoringTableTargetXlsxFilePath)
                # 11. after 5 daily, set daily messages to 0 and wait 18 hours until next round.
                if dailyMessagesSent == maxMessageRound:
                    print("Next Sending round In:")
                    # Timer Function.
                    if lf.timeToNextMessagingRound(nextMessagingRoundInSeconds, exitKey):
                        break
                    dailyMessagesSent = 0

        # 12. add 1 to variable i (go to the next person in the dataFrame).
        i = i + 1

    # Getting our amount of columns in the Data Frame.
    numOfColumns = len(df.columns)
    width = 30
    lf.styleExportedXlsxFile(monitoringTableTargetXlsxFilePath, numOfColumns, width)

    # printing our final updated data frame to the console.
    lf.printDataFrame(df)
    # Closing the browser.
    driver.close()
    print("All messages sent!")

except Exception as e:
    print(f"Exception in main\n{e}")
