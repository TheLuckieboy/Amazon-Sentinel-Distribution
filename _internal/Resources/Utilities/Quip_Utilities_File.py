import time, pyautogui, pyperclip, json, os, sys

from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from Resources.Utilities.Utilities_File import StopFunctionException, check_stop_event

from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException, \
    ElementClickInterceptedException
import pyautogui

# ------------------------------------------- #

gtime = 0.25

# ---------------------------------------- #

def Quip_GetInfo_CellText(driver, row=0, column=0, stop_event=None):
    try:
        for attempt in range(9):
            check_stop_event(stop_event)
            try:
                # Construct the Numbers string
                Numbers = '-cell-' + str(row) + '-' + str(column)

                # Find the element by its class attribute
                input_element = driver.find_element(By.CSS_SELECTOR, '[id^="id-temp"][id$="' + Numbers + '"]')

                # Get the text content of the cell
                cell_text = input_element.text

                # If no exception occurred, return the cell text
                return cell_text

            except NoSuchElementException:
                cells = driver.find_elements(By.CLASS_NAME, 'spreadsheet-cell.react-cell.document-content.first-col')
                FixRow, FixCol = None, None

                for cell in cells:
                    check_stop_event(stop_event)
                    try:
                        cell_id = cell.get_attribute("id")
                        _, _, _, FixRow, FixCol = cell_id.split("-")
                        break
                    except ValueError:
                        print("ValueError")

                if FixRow is not None and FixCol is not None:
                    try:
                        Numbers = '-cell-' + str(FixRow) + '-' + str(FixCol)
                        input_element = driver.find_element(By.CSS_SELECTOR, '[id^="id-temp"][id$="' + Numbers + '"]')
                        input_element.click()

                        presses = int(FixRow) - row

                        if presses < 0:
                            pyautogui.press('down', presses=abs(presses) + 3)
                        elif presses > 0:
                            pyautogui.press('up', presses=presses + 3)

                    except NoSuchElementException:
                        print(f"Element with {Numbers} still not found.")
                else:
                    print("FixRow or FixCol not found, skipping this attempt.")
                    continue

            except UnboundLocalError:
                print("UnboundLocalError: Ensure FixRow and FixCol are defined before usage.")
                return False

        return False

    except (StopFunctionException, ElementClickInterceptedException, StaleElementReferenceException) as e:
        print(f"Exception occurred: {e}")
        return False

def Quip_ClickOn_Cell(driver, row=0, column=0, stop_event=None):
    def ExceptionFunction(driver, row, stop_event=None):
        try:
            # Find the element by its class attribute
            cells = driver.find_elements(By.CLASS_NAME, 'spreadsheet-cell.react-cell.document-content.first-col')
            for cell in cells:
                try:
                    cell_id = cell.get_attribute("id")
                    _, _, _, FixRow, FixCol = cell_id.split("-")
                    break
                except ValueError:
                    pass

            try:
                # Construct the Numbers string
                Numbers = '-cell-' + str(FixRow) + '-' + str(FixCol)

                # Find the element by its class attribute
                input_element = driver.find_element(By.CSS_SELECTOR, '[id^="id-temp"][id$="' + Numbers + '"]')
                input_element.click()

                # Calculate the number of times to press the down or up arrow
                presses = int(FixRow) - row

                # Press the down or up arrow that many times
                if presses < 0:
                    check_stop_event(stop_event)
                    while not stop_event.is_set():
                        pyautogui.press('down', presses=abs(presses)+3)
                        break
                    check_stop_event(stop_event)
                elif presses > 0:
                    check_stop_event(stop_event)
                    while not stop_event.is_set():
                        pyautogui.press('up', presses=presses + 3)
                        break
                    check_stop_event(stop_event)
            except UnboundLocalError:
                pass
        except StopFunctionException:
            return False

    try:
        for attempt in range(3):
            check_stop_event(stop_event)
            try:
                # Construct the Numbers string
                Numbers = '-cell-' + str(row) + '-' + str(column)

                # Find the element by its class attribute
                input_element = driver.find_element(By.CSS_SELECTOR, '[id^="id-temp"][id$="' + Numbers + '"]')
                input_element.click()

                return input_element
            except NoSuchElementException:
                if not ExceptionFunction(driver, row, stop_event=stop_event):
                    return False
            """except ElementClickInterceptedException:
                ExceptionFunction(driver, row, stop_event=stop_event)"""

    except (StopFunctionException, ElementClickInterceptedException, StaleElementReferenceException):
        return False

def Quip_ClickOn_Bucket(driver):
    try:
        # Find the button element with the label "Background Color"
        element = driver.find_element(By.CSS_SELECTOR, 'button[aria-label="Background Color"]')

        # Click on the button
        element.click()
    except NoSuchElementException:
        pyautogui.press('esc')
    except ElementClickInterceptedException:
        return

def Quip_Check_CommandLine(driver, row=0, column=0, stop_event=None):
    try:
        result = Quip_GetInfo_CellText(driver, row, column, stop_event=stop_event)
        if not result:
            return False
        result = str(result).lower()
        if result == "skip":
            print("Skip")
            return result
        elif result == "stopall":
            print("Stopall")
            return result
        else:
            print("Continue")
            return result
    except (StopFunctionException, ElementClickInterceptedException, StaleElementReferenceException):
        return False

def Quip_GetInfo_LegalName(driver, stop_event=None):
    try:
        check_stop_event(stop_event)
        pyautogui.hotkey('ctrl', 'f')
        check_stop_event(stop_event)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1)
        check_stop_event(stop_event)
        pyautogui.press('enter')
        check_stop_event(stop_event)
        pyautogui.press('esc')
        check_stop_event(stop_event)
        pyautogui.press('left', presses=2)
        check_stop_event(stop_event)
        pyautogui.hotkey('ctrl', 'c')
        check_stop_event(stop_event)
        FirstName = str(pyperclip.paste())
        pyautogui.press('right')
        check_stop_event(stop_event)
        pyautogui.hotkey('ctrl', 'c')
        LastName = str(pyperclip.paste())
        check_stop_event(stop_event)
        MixedName = f"{LastName},{FirstName}"
        check_stop_event(stop_event)
        return True, MixedName, FirstName, LastName

    except (StopFunctionException, ElementClickInterceptedException, StaleElementReferenceException):
        return False, False, False, False

# --------------------------------------------------------------------------- #

def Quip_Color_Cells(driver, Color, WorkingRow, Column="0", Row=True, stop_event=None):
    try:
        try:
            pyautogui.press('esc')
            check_stop_event(stop_event)
            #time.sleep(gtime)
            Quip_ClickOn_Cell(driver, WorkingRow, Column, stop_event=stop_event)
            if Row:
                actions = ActionChains(driver)
                actions.key_down(Keys.SHIFT).send_keys(Keys.SPACE).key_up(Keys.SHIFT).perform()
            Quip_ClickOn_Bucket(driver)

            if Color == "None":
                driver.find_element(By.CLASS_NAME, 'color-clear-swatch.button.button-flex.bordered.clickable').click()
                pyautogui.press('down')
                check_stop_event(stop_event)
                time.sleep(gtime)
            else:
                # Capitalize the first letter of the color string
                Color = Color.lower()
                formatted_color = Color.capitalize()

                driver.find_element(By.CSS_SELECTOR, f'div.color-swatch[title="{formatted_color}"]').click()
                pyautogui.press('down')
                check_stop_event(stop_event)
                #time.sleep(gtime)
        except NoSuchElementException:
            pyautogui.press('esc')
    except (StopFunctionException, ElementClickInterceptedException, StaleElementReferenceException):
        return False