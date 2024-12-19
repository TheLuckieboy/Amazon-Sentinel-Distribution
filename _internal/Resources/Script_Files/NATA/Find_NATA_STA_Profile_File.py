from Resources.Utilities.Quip_Utilities_File import Quip_GetInfo_CellText, Quip_ClickOn_Cell, Quip_ClickOn_Bucket, Quip_Check_CommandLine, Quip_GetInfo_LegalName, Quip_Color_Cells
from Resources.Script_Files.NATA.NATACS_Utilities_File import NATA_NameSplitter_LastFirst, NATA_PasteSearchInfo, NATA_WaitFor_Loading, NATA_IsProfile_Present, NATA_Find_Profile, NATA_DeactivateEmployee
from Resources.Utilities.Utilities_File import StopFunctionException, check_stop_event
import time, pyautogui, pyperclip, sys, os

gtime = 0.25

def Find_NATA_STA_Profile(driver, window_handles, WorkingRow, settings, stop_event=None):
    try:
        import json

        def SwithTo_Window(QuipWindow=False, NATACS=False):
            check_stop_event(stop_event)
            if QuipWindow:
                driver.switch_to.window(window_handles[0])
                check_stop_event(stop_event)
                time.sleep(gtime)

            if NATACS:
                driver.switch_to.window(window_handles[1])
                check_stop_event(stop_event)
                time.sleep(gtime)

        check_stop_event(stop_event)
        if WorkingRow <= 0:
            print("Starting Row can not be 1")
            return False

        CommandLineColumn = ord((settings.get("CommandLine Column", "A")).upper()) - ord('A')

        SearchByColumn_FirstName = ord((settings.get("SearchByColumn_FirstName_NATA", "A")).upper()) - ord('A')
        SearchByColumn_LastName = ord((settings.get("SearchByColumn_LastName_NATA", "A")).upper()) - ord('A')
        SearchByFullName, SearchByColumn_FullName = settings.get("SearchByColumn_FullName_NATA", [False, "A"])
        SearchByColumn_FullName = ord(SearchByColumn_FullName.upper()) - ord('A')
        SearchByLogin, SearchByColumn_Login = settings.get("SearchByColumn_Login_NATA", [False, "A"])
        SearchByColumn_Login = ord(SearchByColumn_Login.upper()) - ord('A')

        Quip_Database = settings.get("Quip_Database", True)
        Excel_Database = settings.get("Excel_Database", False)

        if Quip_Database and not Excel_Database:
            def Color_InfoTo_Quip(TooManyProfiles=False, NoProfileFound=False, ProfileFound=False, TooYoung=False):
                # Iterate over settings and perform actions
                settings_to_process = [
                    "TooManyProfiles", "NoProfileFound", "ProfileFound", "TooYoung"
                ]
                for setting_name in settings_to_process:
                    check_stop_event(stop_event)
                    boolean_value, Color = settings.get(setting_name, (False, None))
                    if boolean_value:
                        check_stop_event(stop_event)
                        if setting_name == "TooManyProfiles" and TooManyProfiles:
                            check_stop_event(stop_event)
                            Quip_Color_Cells(driver, Color, WorkingRow, stop_event=stop_event)
                            pyautogui.press('esc')
                        elif setting_name == "NoProfileFound" and NoProfileFound:
                            check_stop_event(stop_event)
                            Quip_Color_Cells(driver, Color, WorkingRow, stop_event=stop_event)
                            pyautogui.press('esc')
                        elif setting_name == "ProfileFound" and ProfileFound:
                            check_stop_event(stop_event)
                            Quip_Color_Cells(driver, Color, WorkingRow, stop_event=stop_event)
                            pyautogui.press('esc')
                        elif setting_name == "TooYoung" and TooYoung:
                            check_stop_event(stop_event)
                            Quip_Color_Cells(driver, Color, WorkingRow, stop_event=stop_event)
                            pyautogui.press('esc')

                # Iterate over settings and perform actions
                settings_to_process = ["*Color a Column* Row Complete_Widget_NATA"]
                for setting_name in settings_to_process:
                    check_stop_event(stop_event)
                    boolean_value, column_name, Color = settings.get(setting_name, (False, "A", None))
                    if boolean_value:
                        check_stop_event(stop_event)
                        column_name = ord(column_name.upper()) - ord('A')
                        Quip_Color_Cells(driver, Color, WorkingRow, Row=False, Column=column_name, stop_event=stop_event)
                        pyautogui.press('esc')

            check_stop_event(stop_event)
            CommandLineResult = Quip_Check_CommandLine(driver, WorkingRow, CommandLineColumn, stop_event=stop_event)
            check_stop_event(stop_event)
            if CommandLineResult == "skip":
                return True
            elif CommandLineResult != "stopall":
                check_stop_event(stop_event)
                SearchByInfo1 = None
                SearchByInfo2 = None

                if SearchByFullName:
                    FullName = Quip_GetInfo_CellText(driver, WorkingRow, SearchByColumn_FullName, stop_event=stop_event)
                    if settings.get("LastNameComma_FirstName_NATA", False):
                        SearchByInfo1, SearchByInfo2 = NATA_NameSplitter_LastFirst(FullName, LastNameComma_FirstName=True,
                                                                             stop_event=stop_event)
                    elif settings.get("FirstNameComma_LastName_NATA", False):
                        SearchByInfo1, SearchByInfo2 = NATA_NameSplitter_LastFirst(FullName, FirstNameComma_LastName=True,
                                                                             stop_event=stop_event)
                    elif settings.get("FirstName_LastName_NATA", False):
                        SearchByInfo1, SearchByInfo2 = NATA_NameSplitter_LastFirst(FullName, FirstName_LastName=True,
                                                                             stop_event=stop_event)
                    else:
                        print("SearchByFullName Interpreter settings invalid")
                        SwithTo_Window(QuipWindow=True)
                        time.sleep(gtime)
                        return False
                
                elif SearchByLogin:
                    SearchByInfo1 = Quip_GetInfo_CellText(driver, WorkingRow, SearchByColumn_Login,
                                                      stop_event=stop_event)
                else:
                    SearchByInfo1 = Quip_GetInfo_CellText(driver, WorkingRow, SearchByColumn_FirstName,
                                                      stop_event=stop_event)
                    SearchByInfo2 = Quip_GetInfo_CellText(driver, WorkingRow, SearchByColumn_LastName, stop_event=stop_event)

                check_stop_event(stop_event)
                print("SearchByInfo: ", SearchByInfo1, SearchByInfo2)

                if SearchByInfo1 is not None and not False:
                    Quip_ClickOn_Cell(driver, WorkingRow - 1, CommandLineColumn, stop_event=stop_event)
                    pyautogui.press('down', presses=3)
                    time.sleep(gtime)

                    # Switch to Skyline
                    SwithTo_Window(NATACS=True)
                    time.sleep(gtime)

                    if NATA_PasteSearchInfo(driver, SearchByInfo1, SearchByInfo2, stop_event=stop_event):
                        check_stop_event(stop_event)
                        time.sleep(gtime)
                        Status1, Result = NATA_WaitFor_Loading(driver, stop_event=stop_event)

                        if Status1:
                            check_stop_event(stop_event)
                            if Result:
                                check_stop_event(stop_event)
                                if NATA_IsProfile_Present(driver):
                                    check_stop_event(stop_event)
                                    DeactivateEmployee = settings.get("DeactivateEmployee", False)
                                    check_stop_event(stop_event)
                                    if DeactivateEmployee:
                                        check_stop_event(stop_event)
                                        Status2, Link = NATA_Find_Profile(driver, SearchByInfo1, SearchByInfo2, stop_event=stop_event)

                                        if Status2:
                                            check_stop_event(stop_event)
                                            if Link != "TooManyProfiles":
                                                check_stop_event(stop_event)
                                                if Link != "NoProfiles":
                                                    check_stop_event(stop_event)
                                                    driver.execute_script("window.open('" + Link + "', '_blank');")
                                                    check_stop_event(stop_event)
                                                    driver.switch_to.window(driver.window_handles[-1])
                                                    check_stop_event(stop_event)

                                                    Reason = settings.get("DeactivationReason", "Terminated")
                                                    ByWho = settings.get("DeactivatedBy", "SentinalAutoScript")

                                                    if NATA_DeactivateEmployee(driver, Reason, ByWho, stop_event=stop_event):
                                                        check_stop_event(stop_event)
                                                        driver.close()
                                                        check_stop_event(stop_event)
                                                        time.sleep(gtime)
                                                        check_stop_event(stop_event)
                                                        SwithTo_Window(QuipWindow=True)
                                                        check_stop_event(stop_event)
                                                        return True

                                                    else:
                                                        print("Killswitch")
                                                        driver.close()
                                                        SwithTo_Window(QuipWindow=True)
                                                        time.sleep(gtime)
                                                        return False
                                                else:
                                                    check_stop_event(stop_event)
                                                    SwithTo_Window(QuipWindow=True)
                                                    check_stop_event(stop_event)
                                                    time.sleep(gtime)
                                                    check_stop_event(stop_event)
                                                    Color_InfoTo_Quip(NoProfileFound=True)
                                                    check_stop_event(stop_event)
                                                    return True

                                            else:
                                                check_stop_event(stop_event)
                                                SwithTo_Window(QuipWindow=True)
                                                check_stop_event(stop_event)
                                                time.sleep(gtime)
                                                check_stop_event(stop_event)
                                                Color_InfoTo_Quip(TooManyProfiles=True)
                                                check_stop_event(stop_event)
                                                return True
                                        else:
                                            print("Status2: ", Status2)
                                            SwithTo_Window(QuipWindow=True)
                                            time.sleep(gtime)
                                            return False
                                    else:
                                        Status, ColorStyle = NATA_Find_Profile(driver, SearchByInfo1=SearchByInfo1, SearchByInfo2=SearchByInfo2, CheckDate="01/01/2024", stop_event=stop_event)
                                        if Status:
                                            check_stop_event(stop_event)
                                            SwithTo_Window(QuipWindow=True)
                                            check_stop_event(stop_event)
                                            time.sleep(gtime)
                                            if ColorStyle == "TooManyProfiles":
                                                Color_InfoTo_Quip(TooManyProfiles=True)
                                                check_stop_event(stop_event)
                                                return True
                                            elif ColorStyle == "NoProfiles":
                                                Color_InfoTo_Quip(NoProfileFound=True)
                                                check_stop_event(stop_event)
                                                return True
                                            elif ColorStyle == "TooYoung":
                                                Color_InfoTo_Quip(TooYoung=True)
                                                check_stop_event(stop_event)
                                                return True
                                            else:
                                                Color_InfoTo_Quip(ProfileFound=True)
                                                check_stop_event(stop_event)
                                                return True
                                        else:
                                            print("Status: ", Status)
                                            SwithTo_Window(QuipWindow=True)
                                            time.sleep(gtime)
                                            return False
                                else:
                                    print("No Profile found")
                                    check_stop_event(stop_event)
                                    SwithTo_Window(QuipWindow=True)
                                    check_stop_event(stop_event)
                                    time.sleep(gtime)
                                    return True

                            else:
                                print("Result: ", Result)
                                SwithTo_Window(QuipWindow=True)
                                time.sleep(gtime)
                                return False
                        else:
                            print("Status: ", Status1)
                            SwithTo_Window(QuipWindow=True)
                            time.sleep(gtime)
                            return False
                    else:
                        print("NATA_PasteName: ", SearchByInfo1," ", SearchByInfo2)
                        SwithTo_Window(QuipWindow=True)
                        time.sleep(gtime)
                        return False
                else:
                    print("FirstName came back with False")
                    SwithTo_Window(QuipWindow=True)
                    time.sleep(gtime)
                    return False
            else:
                print("stop all")
                SwithTo_Window(QuipWindow=True)
                time.sleep(gtime)
                return False

        elif not Quip_Database and Excel_Database:
            print("Does not exsist, Work in Progress")
            SwithTo_Window(QuipWindow=True)
            time.sleep(gtime)
            return False

    except StopFunctionException:
        print("Function stopped by kill switch.")
        return False