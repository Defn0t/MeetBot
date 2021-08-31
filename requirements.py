import selenium.common.exceptions
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from time import sleep


def login():
    url = "https://accounts.google.com/signin/v2/identifier?flowName=GlifWebSignIn&flowEntry=ServiceLogin"
    email = "mihirshah_en20a17_75@dtu.ac.in"
    pas = "gosuckdick"

    xpath = {
        "email_box": '//*[@id="identifierId"]',
        "next_button": '//*[@id="identifierNext"]/div/button',
        "pas_box": '//*[@id="password"]/div[1]/div/div[1]/input',
        "signin_button": '//*[@id="passwordNext"]/div/button',
    }

    options = webdriver.ChromeOptions()
    #  options.add_argument("--mute-audio")
    # Removes navigator.webdriver flag
    # For ChromeDriver version 79.0.3945.16 or over
    options.add_argument('--disable-blink-features=AutomationControlled')
    # pre =1 (allow) and 2 (block)
    # bug exploit fails occasionally: reason unknown
    options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.media_stream_mic": 2,
        "profile.default_content_setting_values.media_stream_camera": 2,
        "profile.default_content_setting_values.geolocation": 2,
        "profile.default_content_setting_values.notifications": 2
    })

    driver = webdriver.Chrome(executable_path='chromedriver.exe', options=options)
    wait = WebDriverWait(driver, 60)
    driver.get(url)

    email_box = wait.until(ec.presence_of_element_located((By.XPATH, xpath["email_box"]))).send_keys(email)
    next_button = wait.until(ec.element_to_be_clickable((By.XPATH, xpath["next_button"]))).click()
    pas_box = wait.until(ec.visibility_of_element_located((By.XPATH,xpath["pas_box"]))).send_keys(pas)
    try:
        signin_button = wait.until(ec.element_to_be_clickable((By.XPATH, xpath["signin_button"]))).click()
    except selenium.common.exceptions.ElementClickInterceptedException:
        pass
    sleep(3)
    print(">Google login complete")
    return driver


def join_again():
    global pause
    global current_class
    code = current_class[1]

    if input(f"You had exited {code} earlier. Do you want to rejoin [Any]? "):
        global store_id
        store_id.clear()
        pause = False


def main():
    main_driver = open_chrome("Profile 7")
    i = 1
    while i > 0:
        call = time_check()
        global current_class
        current_class = call
        if call is None:
            print(">EOD")
            break

        elif call[0] in store_id:
            thread_wait_till_next = Thread(target=wait_till_next, daemon=True)
            thread_join_again = Thread(target=join_again, daemon=True)
            thread_wait_till_next.start()
            thread_join_again.start()
            # Use join method to wait for threads to be completed or else they will be destroyed
            # as they are will be out of scope, since the main function does not wait for their completion
            # it will move forward
            thread_wait_till_next.join()

        elif pause is False:
            new_class = Schedule(call[0], call[1], call[2], call[3], main_driver)
            store_id.append(call[0])
            thread_auto = Thread(target=new_class.exit_meeting_auto, daemon=True)
            thread_input = Thread(target=new_class.exit_meeting_input, daemon=True)
            thread_attendance = Thread(target=new_class.attendance, daemon=True)

            thread_auto.start()
            thread_input.start()
            thread_attendance.start()

            # A class object cannot be deleted until all the threads of that said object have ended
            # Because a class object cannot be destroyed
            # Deleting the object while your thread accesses it in an unordered (non-synchronized, basically)
            # manner is undefined behavior, even if you get lucky.
            thread_auto.join()
            thread_attendance.join()

            del new_class


def wait_till_next():
    global pause
    pause = True
    global current_class
    now = datetime.today()
    exit_time = current_class[3]
    pause_time = exit_time - now
    print(f">Sleeping for {pause_time}")
    pause_time = pause_time.total_seconds()
    while datetime.now() - now < timedelta(seconds=pause_time):
        if pause is False:
            break
    print(">Sleep completed")
    global store_id
    store_id.clear()
    pause = False


# Class V1

class Schedule:

    def __init__(self, code, entry_time, exit_time, driver):
        print(f">new_class created:[{datetime.now()}] code :: {code} :: entry_time :: {entry_time} :: exit_time :: {exit_time} ")
        self.code = code
        self.entry_time = entry_time
        self.exit_time = exit_time
        self.link = self.links[code]
        self.driver = driver
        self.exited = False
        if self.link is None:
            self.link = input(f"Please enter the link for {self.code} :: ")
        self.link = f"'{self.link}'"
        self.join_meet()

    links = {
        "AP": None,
        "EE": "https://meet.google.com/dju-rikc-dqp",
        "MA": "https://meet.google.com/lookup/egu4culb4h?authuser=0&hs=179",
        "CO": "https://meet.google.com/lookup/eflwgh5lju?authuser=0&hs=179",
        "EELAB": None,
        "COLAB": "https://meet.google.com/lookup/fpslolbjbv?authuser=0&hs=179",
        "APLAB": "https://meet.google.com/fha-actr-rqk",
        "MELAB": "https://meet.google.com/lookup/ay6bqb34ol?authuser=0&hs=179",
        "BCOM": 'https://meet.google.com/tom-cwzr-tce'
    }

    xpath = {
        "dismiss_button": '//*[@id="yDmH0d"]/div[3]/div/div[2]/div[3]/div/span',
        "ask_to_join_button": '//*[@id="yDmH0d"]/div[3]/div/div[2]/div[3]/div',
        "join_button": '//*[@id="yDmH0d"]/c-wiz/div/div/div[9]/div[3]/div/div/div[4]/div/div/div[2]/div/div[2]/div/div[1]/div[1]/span/span',
        "show_people": '//*[@id="ow3"]/div[1]/div/div[9]/div[3]/div[10]/div[3]/div[2]/div/div/div[2]/span/button/i[1]'
    }

    def join_meet(self):
        driver = self.driver
        driver.switch_to.window(driver.window_handles[0])
        wait = WebDriverWait(driver, 60)
        driver.execute_script(f"window.open({self.link});")
        tabs = driver.window_handles
        current_tab = tabs[len(tabs) - 1]
        driver.switch_to.window(current_tab)
        dismiss_button = wait.until(ec.visibility_of_element_located((By.XPATH, self.xpath["dismiss_button"])))
        action_1 = ActionChains(driver)
        action_1.move_to_element(dismiss_button).click().perform()

        try:
            sleep(5)
            print(">joining via join button")
            join_button = wait.until(ec.visibility_of_element_located((By.XPATH, self.xpath["join_button"])))
            action_3 = ActionChains(driver)
            action_3.move_to_element(join_button).click().perform()

        except selenium.common.exceptions.TimeoutException:
            print(">failed to locate join button")
            try:
                print(">joining via ask-to-join button")
                ask_button = wait.until(ec.visibility_of_element_located((By.XPATH, self.xpath["ask_to_join_button"]))).click()

            finally:
                pass
        print(">Class joined")

    def roll_number(self, encore):
        encore = encore.lower()
        encore = encore.split(" ")
        temp = ''
        for split in encore:
            if split.isalpha() is False:
                temp += split

        cross = ''
        for i in temp:
            if i.isnumeric():
                cross += i
            elif i == 'o':
                cross += i

        cross = cross.replace("17", '', 1)
        cross = cross.replace("22", '', 1)
        cross = cross.replace('o', '0')
        if len(cross) == 0:
            cross = '0'

        cross = int(cross)
        if cross == 7500:
            cross = 75
        return cross

    def attendance(self):
        run_interval = 2
        driver = self.driver
        wait = WebDriverWait(driver, 12000000)
        attendance_file = f"{self.code}.xlsx"
        wb = load_workbook(attendance_file)
        ws = wb.active

        # Creating a list of members which already exist in the xl file
        attendee_names_xl = []
        attendees = ws.iter_cols(min_row=2, max_row=ws.find_eod_cell, min_col=1, max_col=1)
        # Even though a singular tuple is being returned in this case
        # The generator objects acts ina  way that multiple tuples are
        # being returned hence we must unpack the first tuple and then access it
        for tuple_1 in attendees:
            attendees = tuple_1

        # Checking for previously existing names in xl file
        for attendee in attendees:
            attendee_names_xl.append(attendee.value)

        # Finding the date column/ Creating one if it doesnt exist already
        max_col = ws.max_column
        dates = ws.iter_rows(min_row=1, max_row=1, min_col=2, max_col=max_col)
        for tuple_1 in dates:
            dates = tuple_1
        date_cell = None

        for date in dates:
            if date.value == self.entry_time:
                date_cell = date
                break

        if date_cell is None:
            date_cell = ws.cell(row=1, column=max_col+1, value=self.entry_time)
        main_col = date_cell.column

        print("")
        print(">", date_cell)
        print(">", date_cell.value)
        print(">Opening attendee list")
        join_button = wait.until(ec.visibility_of_element_located((By.XPATH, self.xpath["show_people"])))
        action_3 = ActionChains(driver)
        action_3.move_to_element(join_button).click().perform()
        sleep(5)

        while self.exited is False:
            start_check = datetime.now()
            print(f">[{datetime.today()}] Checking for attendees")
            names = []
            html_text = driver.page_source
            soup = BeautifulSoup(html_text, 'lxml')
            list_ = soup.find("div", {"role": "list"})
            items = list_.find_all("div", {"role": "listitem"})
            for item in items:
                text = item.text
                names.append(text)

            # adding new attendees to the xl file
            for name in names:
                if name not in attendee_names_xl:
                    new_attendee = ws.cell(row=ws.find_eod_cell + 1, column=1, value=name)
                    attendee_names_xl.append(name)

            # Refreshing the tuple containing cell objects after adding new cells in xl file
            attendees = ws.iter_cols(min_row=2, max_row=ws.find_eod_cell, min_col=1, max_col=1)
            for tuple_1 in attendees:
                attendees = tuple_1

            # Adding time for each attendee
            for attendee in attendees:
                if attendee.value in names:

                    time_add = ws.cell(row=attendee.row, column=main_col)
                    time_elapsed = datetime.now() - start_check
                    time_elapsed = time_elapsed.total_seconds()
                    time_elapsed = time_elapsed/60

                    if time_add.value is None:
                        time_add.value = run_interval + time_elapsed
                    else:
                        time_add.value += run_interval + time_elapsed
                else:
                    pass

            # So that the thread terminates once the class has ended even if its is sleeping for 'run_interval'
            now = datetime.today()
            while datetime.today() - now < timedelta(minutes=run_interval):
                if self.exited is True:
                    break
                else:
                    pass

        # Assigning roll numbers to meet members
        for cell in range(2, ws.find_eod_cell + 1):
            encore = ws['A' + str(cell)].value
            ws['B' + str(cell)].value = self.roll_number(encore)
            print(ws['B' + str(cell)].value)

        wb.save(f"{self.code}.xlsx")
        wb.close()

    def exit_meeting_input(self):
        while self.exited is False:

            if input("Exit now ?[Any]: "):
                self.driver.close()
                print(">Exited class")
                self.exited = True
            else:
                pass

    def exit_meeting_auto(self):
        now = datetime.today()
        until_exit = self.exit_time - now + timedelta(minutes=auto_exit_after)
        print(f">[{now.time()}] Auto-exit in {until_exit}")
        while datetime.today() - now < until_exit:
            if self.exited is True:
                return None
            else:
                pass
        self.driver.close()
        print(">Exited class")
        self.exited = True

    def __del__(self):
        print(f">Class object killed: [{datetime.now()}] id :: {self.code} :: entry_time :: {self.entry_time} :: exit_time :: {self.exit_time}")

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from time import sleep
from bs4 import BeautifulSoup
import lxml
xpath = {
    "search_box": '//*[@id="side"]/div[1]/div/label/div/div[2]',
    "chat": '//*[@id="pane-side"]/div[1]/div/div/div[10]/div/div/div[2]/div[1]/div[1]'
}
options = webdriver.ChromeOptions()
options.add_argument("--user-data-dir=C:/Users/admin/AppData/Local/Google/Chrome/User Data")
options.add_argument('--profile-directory=Profile 7')
options.add_argument('--disable-blink-features=AutomationControlled')
options.add_experimental_option("prefs", {
    "profile.default_content_setting_values.media_stream_mic": 2,
    "profile.default_content_setting_values.media_stream_camera": 2,
    "profile.default_content_setting_values.geolocation": 2,
    "profile.default_content_setting_values.notifications": 2
})


driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 60)
driver.get("https://web.whatsapp.com/")
search_box = wait.until(ec.visibility_of_element_located((By.XPATH, xpath["search_box"]))).send_keys("A17 Unofficial")
sleep(2)
search_button = wait.until(ec.visibility_of_element_located((By.CSS_SELECTOR, "span[title= 'A17 Unofficial']"))).click()
sleep(2)
html_text = driver.page_source
chat = BeautifulSoup(html_text, 'lxml')
chat_window = chat.find('div', {"id": "main"})
table = chat_window.find("div", {"aria-label": "Message list. Press right arrow key on a message to open message context menu."})
for message in table:
    print(message)


def exit_meeting(self):
    now = datetime.today()
    until_exit = self.exit_time - now + timedelta(minutes=3)
    print(f">[{now.time()}] Auto-exit in {until_exit}")

    while datetime.today() - now < until_exit:
        if self.exited is True:
            break
        else:
            pass

    self.exited = True
    while self.attendee_list_open is True:
        pass

    self.driver.close()
    print(">Exited class")

# menu v1
def menu(class_):
    global auto_exit_after
    while True:

        if class_.exit_time + datetime.timedelta(minutes=auto_exit_after) < datetime.datetime.now():
            print(">Killing class object")
            del class_
            break
        print("""[Exit]: Exit current class.
        [Join]: Rejoin current class.
        [Refresh]: Exit and rejoin the current class.
        """)
        x = pytimedinput.timedInput("Please enter any of the above options:: ", timeOut=30)
        input_bool = x[1]
        if input_bool is True:
            print(">No input received!")

        elif input_bool is False:
            user_input = x[0]
            user_input.lower()

            if user_input == "exit":
                if class_.exited is False:
                    class_.exited = True
                    if class_.exit_time < datetime.datetime.now():
                        print(">Killing class object")
                        del class_
                        break
                elif class_.exited is True:
                    print(">You have already exited this class.")

            elif user_input == "Join":
                if class_.exited is True:
                    class_.exited = False
                    class_.joint_meet()
                elif class_.exited is False:
                    print(">You are already in the class")
                    print(">Enter 'refresh' to exit and join again.")

            elif user_input == "refresh":
                print(">Exiting class")
                class_.exited = True
                sleep(5)
                print(">Joining meet again")
                class_.joint_meet()

    print(">Class over!")


