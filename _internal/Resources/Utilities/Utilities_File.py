try:
    class StopFunctionException(Exception):
        pass


    def check_stop_event(stop_event=None):
        if stop_event.is_set():
            raise StopFunctionException


    def Plugins(settings):
        Plugins = []
        Cardholder_Verification_Plugin = settings.get("Cardholder_Verification_Plugin", False)
        Plugins.append(Cardholder_Verification_Plugin)
        Plugins.append("Cardholder Verification")
        Plugins.append("This script retrieves a list of employee EIDs or Logins from your selected database, "
                       "searches each profile in the Cardholder Management System, and copies the requested "
                       "information back to the database. The available options streamline the data transfer process, "
                       "making it easier to complete your tasks efficiently")

        Skyline_Terminate_AA_Plugin = settings.get("Skyline_Terminate_AA_Plugin", False)
        Plugins.append(Skyline_Terminate_AA_Plugin)
        Plugins.append("Skyline Terminate AA")
        Plugins.append("Message")

        PreferredNames_To_LegalNames_Plugin = False
        Plugins.append(PreferredNames_To_LegalNames_Plugin)
        Plugins.append("Placeholder_Name1")
        Plugins.append("Message")

        Quip_ClearRowColor_Plugin = False
        Plugins.append(Quip_ClearRowColor_Plugin)
        Plugins.append("Placeholder_Name2")
        Plugins.append("Message")

        Find_NATA_Profile = True
        Plugins.append(Find_NATA_Profile)
        Plugins.append("Find NATA Profile")
        Plugins.append("Message")

        return Plugins


    # Scripts and widgets being used, aka Plugins

    def Plugin_Widget_Setup(FunctionsGUI, widget, settings, Save_Widget_Settings, grid_layout, Script_Widgets, index1):
        if index1 == 0:
            try:
                from Resources.Script_Files.Cardholder.Cardholder_Verification_File import Widget_Setup
                Widget_Setup(FunctionsGUI, widget, settings, Save_Widget_Settings, grid_layout, Script_Widgets)
            except ImportError:
                print("Import Failed, Cardholder_Verification_File could not be found, Contact Support if needed")
            except Exception as E:
                print(E)

        elif index1 == 3:
            widget = widget
            Script_Widgets.append(widget)

        elif index1 == 12:
            widget = widget
            Script_Widgets.append(widget)

        return widget


    # Creates the Widgets for the Scripts

    def Script_Launcher(get_next_index, function_mapping, index1):
        if index1 == 0:
            try:
                from Resources.Script_Files.Cardholder.Cardholder_Verification_File import Cardholder_Verification
                next_index = get_next_index(function_mapping)
                function_mapping[next_index] = Cardholder_Verification
            except ImportError:
                print("Import Failed, Cardholder_Verification_File could not be found, Contact Support if needed")
            except Exception as E:
                print(E)

        elif index1 == 3:
            try:
                from Resources.Script_Files.Skyline.Skyline_Terminate_AA_File import Skyline_Terminate_AA
                next_index = get_next_index(function_mapping)
                function_mapping[next_index] = Skyline_Terminate_AA
            except ImportError:
                print("Import Failed, Skyline_Terminate_AA_File could not be found, Contact Support if needed")
            except Exception as E:
                print(E)

        elif index1 == 12:
            try:
                from Resources.Script_Files.NATA.Find_NATA_STA_Profile_File import Find_NATA_STA_Profile
                next_index = get_next_index(function_mapping)
                function_mapping[next_index] = Find_NATA_STA_Profile
            except ImportError:
                print("Import Failed, NATACS_Terminate_AA_File could not be found, Contact Support if needed")
            except Exception as E:
                print(E)


    # Attaches the Script to the selected function

    def Open_Browser(FunctionsGUI, Plugins, function_combobox, Chrome=False, Firefox=False, Edge=False):
        from selenium import webdriver
        import threading, time
        gtime = 0.25

        driver_is_Open = True

        if Chrome:
            # Specify the directory where Chrome will store user data
            user_data_dir = "/ChromeData"

            # Initialize Chrome options
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--start-maximized")
            chrome_options.add_argument("--user-data-dir=" + user_data_dir)

            # Create a new instance of the Chrome driver with the specified options
            driver = webdriver.Chrome(options=chrome_options)

        if Firefox:
            # Specify the directory where Firefox will store user data
            user_data_dir = "/FirefoxData"

            # Initialize Firefox options
            firefox_options = webdriver.FirefoxOptions()
            firefox_options.add_argument("--start-maximized")
            firefox_options.add_argument(user_data_dir)

            # Create a new instance of the Firefox driver with the specified options
            driver = webdriver.Firefox(options=firefox_options)

        if Edge:
            # Specify the directory where Firefox will store user data
            user_data_dir = "/EdgeData"

            # Initialize Firefox options
            edge_options = webdriver.EdgeOptions()
            edge_options.add_argument("--start-maximized")
            edge_options.add_argument(user_data_dir)

            # Create a new instance of the Firefox driver with the specified options
            driver = webdriver.Edge(options=edge_options)

        # Store the window handles in a list
        window_handles = []

        # Navigate to a URL
        driver.get("https://quip-amazon.com/all")
        original_window = driver.current_window_handle
        window_handles.append(original_window)
        time.sleep(gtime)

        if (Plugins[0] and function_combobox.currentText() == Plugins[1]):
            driver.execute_script("window.open('" + "https://cm.gso.amazon.dev/" + "', '_blank');")
            window_handles.append(driver.window_handles[-1])
            time.sleep(gtime)

        if (Plugins[3] and function_combobox.currentText() == Plugins[4]):
            driver.execute_script(
                "window.open('" + "https://lossprevention.amazon.com/air?end=&start=&status=All" + "', '_blank');")
            window_handles.append(driver.window_handles[-1])
            time.sleep(gtime)

        if (Plugins[6] and function_combobox.currentText() == Plugins[7]):
            driver.execute_script("window.open('" + "https://quip-amazon.com/all" + "', '_blank');")
            window_handles.append(driver.window_handles[-1])
            time.sleep(gtime)

        if (Plugins[9] and function_combobox.currentText() == Plugins[10]):
            driver.execute_script(
                "window.open('" + "https://lossprevention.amazon.com/air?end=&start=&status=All" + "', '_blank');")
            window_handles.append(driver.window_handles[-1])
            time.sleep(gtime)

        if (Plugins[12] and function_combobox.currentText() == Plugins[13]):
            driver.execute_script("window.open('" + "https://info.natacs.aero" + "', '_blank');")
            window_handles.append(driver.window_handles[-1])
            time.sleep(gtime)

        # Switch to original_window
        driver.switch_to.window(original_window)

        # Start the thread to periodically check if the driver is open
        check_driver_thread = threading.Thread(target=FunctionsGUI.check_driver_status)
        check_driver_thread.daemon = True
        check_driver_thread.start()

        return driver, window_handles, driver_is_Open


    def Help_GUI(FunctionsGUI, settings, Save_Widget_Settings):
        from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout, QHBoxLayout, QSpacerItem, QSizePolicy
        from PyQt5.QtGui import QFont

        HelpPage_Widget = QWidget()
        section_layout = QVBoxLayout(HelpPage_Widget)
        section_layout.setContentsMargins(24, 8, 24, 8)
        section_layout.setSpacing(8)
        Widgets = []

        LabelText = QLabel(
            "The Help page is a work in progress, please contact support or Michael Luckie, wmichluc, for any help needed")
        LabelText.setFont(QFont("Sans-serif", 22, weight=QFont.Bold))
        LabelText.setWordWrap(True)
        section_layout.addWidget(LabelText)

        return HelpPage_Widget


    # Help GUI

    def Extra_GUI(FunctionsGUI, settings, Save_Widget_Settings):
        from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout, QHBoxLayout, QSpacerItem, QSizePolicy
        from PyQt5.QtGui import QFont

        Plugin_Widget = QWidget()
        section_layout = QVBoxLayout(Plugin_Widget)
        section_layout.setContentsMargins(8, 8, 8, 8)
        section_layout.setSpacing(8)
        Widgets = []

        LabelText = QLabel("Script Plugins:")
        LabelText.setFont(QFont("Sans-serif", 12, weight=QFont.Bold))
        section_layout.addWidget(LabelText)

        Cardholder_Verification_Plugin = FunctionsGUI.Widget_Creator("Cardholder Verification", Font_Size=8,
                                                                     Checkmark=True)
        Cardholder_Verification_Plugin.setObjectName("Cardholder_Verification_Plugin")
        Widgets.append(Cardholder_Verification_Plugin)

        Find_NATA_Profile_Plugin = FunctionsGUI.Widget_Creator("Find NATA Profile", Font_Size=8, Checkmark=True)
        Find_NATA_Profile_Plugin.setObjectName("Find_NATA_Profile_Plugin")
        Widgets.append(Find_NATA_Profile_Plugin)

        for widget in Widgets:
            section_layout.addWidget(widget)

            if widget.objectName() == "Cardholder_Verification_Plugin":
                Cardholder_Verification_Plugin = widget
                Cardholder_Verification_Plugin.checkmark.setChecked(
                    settings.get("Cardholder_Verification_Plugin", False))

            if widget.objectName() == "Find_NATA_Profile_Plugin":
                Find_NATA_Profile_Plugin = widget
                Find_NATA_Profile_Plugin.checkmark.setChecked(settings.get("Find_NATA_Profile_Plugin", False))

            widget.checkmark.stateChanged.connect(lambda: Save_Widget_Settings(Widgets))
            widget.checkmark.stateChanged.connect(lambda: FunctionsGUI.PluginEvent())

        section_layout.addItem(QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))
        HLayout = QHBoxLayout()
        HLayout.addLayout(section_layout)
        HLayout.addItem(QSpacerItem(0, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))

        return Plugin_Widget
    # Extra Tab

except Exception as e:
    print(e)