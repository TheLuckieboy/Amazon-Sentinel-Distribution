import time, pyautogui, pyperclip, datetime

from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, TimeoutException, UnexpectedAlertPresentException, InvalidSessionIdException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from Resources.Utilities.Utilities_File import StopFunctionException, check_stop_event

gtime = 0.25

"""except StopFunctionException:
    print("Function stopped by kill switch.")
    return False"""

def NATA_NameSplitter_LastFirst(MixedName, LastNameComma_FirstName=False, FirstNameComma_LastName=False, FirstName_LastName=False, stop_event=None):
    try:
        new_first_name = None
        new_last_name = None

        if LastNameComma_FirstName:
            # Split the mixed name into parts
            parts = MixedName.split(',')
            check_stop_event(stop_event)

            # Extract the first name and remove leading/trailing spaces
            first_name = parts[1].strip().split()[0]
            check_stop_event(stop_event)

            # Extract the last name and remove leading/trailing spaces
            last_name = parts[0].strip()
            check_stop_event(stop_event)

            # Rearrange the parts and remove leading/trailing spaces
            new_first_name = f"{first_name}"
            new_last_name = f"{last_name}"
            check_stop_event(stop_event)
        elif FirstNameComma_LastName:
            # Split the mixed name into parts
            parts = MixedName.split(',')
            check_stop_event(stop_event)

            # Extract the first name and remove leading/trailing spaces
            first_name = parts[0].strip().split()[1]
            check_stop_event(stop_event)

            # Extract the last name and remove leading/trailing spaces
            last_name = parts[1].strip()
            check_stop_event(stop_event)

            # Rearrange the parts and remove leading/trailing spaces
            new_first_name = f"{first_name}"
            new_last_name = f"{last_name}"
            check_stop_event(stop_event)
        elif FirstName_LastName:
            # Split the mixed name into parts using spaces
            parts = MixedName.split()
            check_stop_event(stop_event)

            # Extract the first name (first element)
            first_name = parts[0].strip()
            check_stop_event(stop_event)

            # Combine the remaining parts as the last name
            last_name = " ".join(parts[1:]).strip()
            check_stop_event(stop_event)

            # Set the new first name and last name
            new_first_name = first_name
            new_last_name = last_name
            check_stop_event(stop_event)

        return new_first_name, new_last_name
    except AttributeError:
        print("AttributeError")
        return False, False

    except StopFunctionException:
        print("Function stopped by kill switch.")
        return False, False

def NATA_PasteSearchInfo(driver, SearchByInfo1, SearchByInfo2=None, stop_event=None):
    try:
        TabsElements = driver.find_element(By.CLASS_NAME, "ant-tabs-nav-list")
        SearchOptionsTab = TabsElements.find_element(By.XPATH, './/div[@data-node-key="1"]')
        check_stop_event(stop_event)
        SearchOptionsTab.click()

        if SearchByInfo2:
            FirstNameElement = driver.find_element(By.ID, 'first')
            LastNameElement = driver.find_element(By.ID, 'last')

            # Triple-click on the input element
            actions = ActionChains(driver)
            check_stop_event(stop_event)
            actions.click(FirstNameElement).click(FirstNameElement).click(FirstNameElement).perform()
            check_stop_event(stop_event)
            FirstNameElement.send_keys(SearchByInfo1)

            # Triple-click on the input element
            actions = ActionChains(driver)
            check_stop_event(stop_event)
            actions.click(LastNameElement).click(LastNameElement).click(LastNameElement).perform()
            check_stop_event(stop_event)
            LastNameElement.send_keys(SearchByInfo2)
            check_stop_event(stop_event)
        else:
            EmployeeNumberElement = driver.find_element(By.ID, 'compEmpNum')

            # Triple-click on the input element
            actions = ActionChains(driver)
            check_stop_event(stop_event)
            actions.click(EmployeeNumberElement).click(EmployeeNumberElement).click(EmployeeNumberElement).perform()
            check_stop_event(stop_event)
            EmployeeNumberElement.send_keys(SearchByInfo1)

        ButtonElements = driver.find_elements(By.CLASS_NAME, 'ant-btn.css-1wazalj.ant-btn-primary')

        for button in ButtonElements:
            check_stop_event(stop_event)
            span_elements = button.find_elements(By.TAG_NAME, 'span')
            for span in span_elements:
                if span.text.strip() == "Search":
                    check_stop_event(stop_event)
                    button.click()
                    return True

    except StopFunctionException:
        print("Function stopped by kill switch.")
        return False

def NATA_WaitFor_Loading(driver, FindProfile=True, stop_event=None):
    try:
        if FindProfile:
            try:
                # Wait until the element with class "modal-backdrop.fade.show" is not present
                WebDriverWait(driver, 10).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, 'ant-spin.ant-spin-spinning.ant-spin-show-text.css-1wazalj'))
                )
                # Element is not present, return True
                time.sleep(1)
                check_stop_event(stop_event)
                return True, True
            except TimeoutException:
                # Handle the case where the element is still present after the timeout
                check_stop_event(stop_event)
                return True, False
        else:
            try:
                # Wait until the element with class "modal-backdrop.fade.show" is not present
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'DeactivationReasonInput'))
                )
                # Element is not present, return True
                time.sleep(1)
                check_stop_event(stop_event)
                return True, True
            except TimeoutException:
                # Handle the case where the element is still present after the timeout
                check_stop_event(stop_event)
                return True, False

    except StopFunctionException:
        print("Function stopped by kill switch.")
        return False, False

def NATA_IsProfile_Present(driver):
    try:
        Temp = driver.find_element(By.CLASS_NAME, 'ant-empty-image')
        return False
    except NoSuchElementException:
        return True

def NATA_DeactivateEmployee(driver, Reason, ByWho, stop_event=None):
    try:
        DeactivateEmployee_button = driver.find_element(By.XPATH, '//input[@value="Deactivate Employee"]')
        check_stop_event(stop_event)
        DeactivateEmployee_button.click()
        time.sleep(gtime)

        NATA_WaitFor_Loading(driver, FindProfile=False, stop_event=stop_event)
        check_stop_event(stop_event)
        time.sleep(gtime)

        DeactivateReasonElement = driver.find_element(By.ID, 'DeactivationReasonInput')
        check_stop_event(stop_event)
        DeactivateReasonElement.send_keys(Reason)
        check_stop_event(stop_event)
        time.sleep(gtime)

        DeactivateByElement = driver.find_element(By.ID, 'DeactivationByInput')
        check_stop_event(stop_event)
        DeactivateByElement.send_keys(ByWho)
        check_stop_event(stop_event)
        time.sleep(gtime)

        Deactivate_button = driver.find_element(By.XPATH, '//input[@value="Deactivate"]')
        check_stop_event(stop_event)
        Deactivate_button.click()
        check_stop_event(stop_event)
        time.sleep(gtime)
        return True

    except StopFunctionException:
        print("Function stopped by kill switch.")
        return False
    
def NATA_Find_Profile(driver, SearchByInfo1, SearchByInfo2=None, CheckDate="01/01/2024", stop_event=None):
    try:
        Links = []
        TooManyProfiles = "TooManyProfiles"
        NoProfiles = "NoProfiles"
        TooYoung = "TooYoung"
        RowElements = driver.find_elements(By.CLASS_NAME, 'ant-table-row.ant-table-row-level-0')

        ListProfiles = []

        for Row in RowElements:
            check_stop_event(stop_event)
            TableCellElements = Row.find_elements(By.CLASS_NAME, 'ant-table-cell')

            DataPoint1 = False
            DataPoint2 = False
            DataPoint3 = False

            if SearchByInfo2:
                check_stop_event(stop_event)
                if TableCellElements[1].text.lower() == SearchByInfo1.lower():
                    check_stop_event(stop_event)
                    DataPoint1 = True
                    check_stop_event(stop_event)
                if TableCellElements[3].text.lower() == SearchByInfo2.lower():
                    check_stop_event(stop_event)
                    DataPoint2 = True
                    check_stop_event(stop_event)
                target_date = datetime.datetime.strptime(CheckDate, "%m/%d/%Y")
                cell_date = datetime.datetime.strptime(TableCellElements[8].text, "%m/%d/%Y")
                if cell_date >= target_date:
                    check_stop_event(stop_event)
                    DataPoint3 = True
                    check_stop_event(stop_event)

                if (DataPoint1 and DataPoint2):
                    check_stop_event(stop_event)
                    ListProfiles.append((DataPoint3, TableCellElements[5].find_element(By.TAG_NAME, 'span').get_attribute('textContent'), TableCellElements[7].text))

            else:
                pass
                """if cell_text.lower() == SearchByInfo1.lower():
                    check_stop_event(stop_event)
                    DataPoint1 = True
                    DataPoint2 = True
                else:
                    try:
                        target_date = datetime.datetime.strptime("01/01/2024", "%m/%d/%Y")
                        cell_date = datetime.datetime.strptime(cell_text, "%m/%d/%Y")
                        if cell_date <= target_date:
                            check_stop_event(stop_event)
                            DataPoint3 = True
                            check_stop_event(stop_event)
                    except ValueError:
                        pass"""
                
        NumProfiles = len(ListProfiles)
        print(ListProfiles)
        print(NumProfiles)
        if NumProfiles == 0:
            check_stop_event(stop_event)
            return True, NoProfiles
        elif NumProfiles > 1:
            ConsolidatedProfiles = []
            A = 0

            while A < len(ListProfiles):
                check_stop_event(stop_event)
                current_profile = (ListProfiles[A][1], ListProfiles[A][2])
                check_stop_event(stop_event)
                if current_profile not in ConsolidatedProfiles:
                    check_stop_event(stop_event)
                    ConsolidatedProfiles.append(current_profile)
                check_stop_event(stop_event)
                A += 1

            if len(ConsolidatedProfiles) > 1:
                check_stop_event(stop_event)
                return True, TooManyProfiles
            else:
                A = 0
                for _ in range(NumProfiles):
                    check_stop_event(stop_event)
                    if ListProfiles[A][0]:
                        check_stop_event(stop_event)
                        return True, TooYoung
                    check_stop_event(stop_event)
                    A += 1
                return True, "Blank"
        else:
            A = 0
            for _ in range(NumProfiles):
                check_stop_event(stop_event)
                if ListProfiles[A][0]:
                    check_stop_event(stop_event)
                    print("TooYoung")
                    return True, TooYoung
                check_stop_event(stop_event)
                A += 1
            print("Blank")
            return True, "Blank"

    except StopFunctionException:
        print("Function stopped by kill switch.")
        return False, None