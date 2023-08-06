from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as CS
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import csv
import time
from selenium.webdriver.common.by import By
from tkinter import filedialog
import getpass
import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
from pathlib import Path
import logging
import threading
import time
import platform
import os

class App:
    def __init__(self, root):
        self.totalContactsTillNow = 0
        self.failedContacts = 0
        self.successfulContacts = 0
        self.currentName = ""
        self.messageStatus = ""
        self.progressBarValue = 0
        self.totalNoContacts = 0
        self.imagePath = ""
        self.contactsFilePath = ""
        self.messageFilePath = ""

        self.root = root
        self.root.title("WhatsApp Bulk Messenger")

        # Dictionary to map display values to code values
        self.display_to_code = {"Yes": "Y", "No": "N"}

        # Calculate the center position of the screen
        window_width = 570
        window_height = 440
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # Set the window size and position
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.attributes("-topmost", True)

        entry_width = 20
        combobox_width = entry_width - 2
        file_selection_frame_width = entry_width - 1
        
        bigfont = Font(family="Helvetica",size=14)
        self.root.option_add("*Font", bigfont)
        file_selection_label_font = ("Helvetica", 14)
        # Set the style for the Combobox elements (options)

        # Create Labels
        self.browser_name_label = tk.Label(self.root, text="Select Browser:")
        self.contact_label = tk.Label(self.root, text="Select contacts file (.csv):")
        self.message_label = tk.Label(self.root, text="Select message  (.txt):")
        self.image_available_label = tk.Label(self.root, text="Do you want to send image?")
        self.image_name_label = tk.Label(self.root, text="Select image:")
        self.are_contacts_with_country_code_label = tk.Label(self.root, text="Are Contacts with Country Code?")
        self.country_code_label = tk.Label(self.root, text="Country Code:")
        self.waiting_time_label = tk.Label(self.root, text="Waiting Time (in seconds):")
        
        self.contacts_file_selected_label = tk.Label(self.root, text="")
        self.message_file_selected_label = tk.Label(self.root, text="")
        # Create Label for the image message
        # self.image_message_label = tk.Label(root, text="Please copy the image to send", fg="red")


        # Create Dropdown (Combobox) for browser name
        self.browser_name_var = tk.StringVar()
        self.browser_name_combobox = ttk.Combobox(
            root, textvariable=self.browser_name_var, values=["Chrome", "Firefox", "Edge"], width=combobox_width
        )
        self.browser_name_combobox.set("Chrome")  # Set the default display value

        # Create Dropdown (Combobox) for "Is Image Available"
        self.image_available_var = tk.StringVar()
        self.image_available_combobox = ttk.Combobox(root, textvariable=self.image_available_var, values=["Yes", "No"], width=combobox_width)
        self.image_available_combobox.set("No")  # Set the default display value to "No"

        # self.image_name_entry = tk.Entry(root, state=tk.DISABLED, width=entry_width)
        # image file selection frame
        self.image_file_selection_frame = tk.Frame(self.root)
        self.image_file_button = tk.Button(self.image_file_selection_frame, text="Select", command=self.open_image, 
                                           state=tk.DISABLED, width=int(file_selection_frame_width/3), font=file_selection_label_font)
        self.image_file_selected_label = tk.Label(self.image_file_selection_frame, width=int(file_selection_frame_width*2/3))

        # Create Dropdown (Combobox) for "Are Contacts with Country Code?"
        self.are_contacts_with_country_code_var = tk.StringVar()
        self.are_contacts_with_country_code_combobox = ttk.Combobox(
            root, textvariable=self.are_contacts_with_country_code_var, values=["Yes", "No"], width=combobox_width
        )
        self.are_contacts_with_country_code_combobox.set("Yes")  # Set the default display value

        # Create Entry widgets
        self.country_code_entry = tk.Entry(root, state=tk.DISABLED, width=entry_width)  # Disabled by default
        self.waiting_time_entry = tk.Entry(root, width=entry_width)  # Waiting time entry
        self.waiting_time_entry.insert(0, "6")

        # Contacts file selection frame
        self.contacts_file_selection_frame = tk.Frame(self.root)
        self.contacts_file_button = tk.Button(self.contacts_file_selection_frame, text="Select", command=self.open_contacts_file
                                              , width=int(file_selection_frame_width/3))
        self.contacts_file_selected_label = tk.Label(self.contacts_file_selection_frame, width=int(file_selection_frame_width*2/3), font=file_selection_label_font)

        # Message file selection frame
        self.message_file_selection_frame = tk.Frame(self.root)
        self.message_file_button = tk.Button(self.message_file_selection_frame, text="Select", command=self.open_message_file
                                             , width=int(file_selection_frame_width/3))
        self.message_file_selected_label = tk.Label(self.message_file_selection_frame, width=int(file_selection_frame_width*2/3), font=file_selection_label_font)

        # Create Submit Button
        self.submit_button = tk.Button(root, text="Submit", command=self.submit_form)

        self.empty_input_label = tk.Label(root)

        # Bind the <Return> event to the submit button
        self.root.bind("<Return>", lambda event: self.submit_form())

        # Bind the callback functions to update the country code entry's state, image message visibility, and "Is Image Available" dropdown
        self.are_contacts_with_country_code_var.trace("w", lambda *args: self.update_country_code_entry(self, *args))
        self.image_available_var.trace("w", lambda *args: self.update_image_message(self, *args))
        self.browser_name_var.trace("w", lambda *args: self.update_is_image_available_dropdown(self, *args))

        # Grid Layout
        self.browser_name_label.grid(row=0, column=0, padx=10, pady=5)
        self.browser_name_combobox.grid(row=0, column=1, padx=10, pady=5)

        self.image_available_label.grid(row=1, column=0, padx=10, pady=5)
        self.image_available_combobox.grid(row=1, column=1, padx=10, pady=5)

        self.image_name_label.grid(row=2, column=0, padx=10, pady=5)
        self.image_file_button.pack(side=tk.LEFT)
        self.image_file_selected_label.pack(side=tk.RIGHT)
        self.image_file_selection_frame.grid(row=2, column=1, padx=10, pady=5)

        self.are_contacts_with_country_code_label.grid(row=3, column=0, padx=10, pady=5)
        self.are_contacts_with_country_code_combobox.grid(row=3, column=1, padx=10, pady=5)

        self.country_code_label.grid(row=4, column=0, padx=10, pady=5)
        self.country_code_entry.grid(row=4, column=1, padx=10, pady=5)

        self.waiting_time_label.grid(row=5, column=0, padx=10, pady=5)
        self.waiting_time_entry.grid(row=5, column=1, padx=10, pady=5)

        self.contact_label.grid(row=6, column=0, padx=10, pady=5)
        self.contacts_file_button.pack(side=tk.LEFT)
        self.contacts_file_selected_label.pack(side=tk.LEFT)
        self.contacts_file_selection_frame.grid(row=6, column=1, padx=10, pady=5)

        self.message_label.grid(row=7, column=0, padx=10, pady=5)
        self.message_file_button.pack(side=tk.LEFT)
        self.message_file_selected_label.pack(side=tk.RIGHT)
        self.message_file_selection_frame.grid(row=7, column=1, padx=10, pady=5)

        self.empty_input_label.grid(row=8, column=0, columnspan=2, padx=10, pady=5)
        self.empty_input_label.grid_forget()

        self.submit_button.grid(row=9, column=0, columnspan=2, padx=10, pady=10)

    def open_image(self):
        imagePath = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.gif;*.jfif")])
        if imagePath:
            self.imagePath = imagePath
            fileName = os.path.basename(self.imagePath)
            self.image_file_selected_label.config(text=fileName)
    
    def open_contacts_file(self):
        contactsFilePath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if contactsFilePath:
            self.contactsFilePath = contactsFilePath
            fileName = os.path.basename(self.contactsFilePath)
            self.contacts_file_selected_label.config(text=fileName)
    
    def open_message_file(self):
        messageFilePath = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if messageFilePath:
            self.messageFilePath = messageFilePath
            fileName = os.path.basename(self.messageFilePath)
            self.message_file_selected_label.config(text=fileName)

    def close_window(self):
        self.root.destroy()
    
    def update_message_to_contact(self):
        self.message_to_contact_label.config(text=f"Currently sending message to {self.currentName}.")

    def update_progress_bar(self):
        self.progress['value'] = self.progressBarValue
        self.root.update_idletasks()

    def update_total_no_contacts(self):
        self.total_no_contacts_label.config(text=f"Total no of contacts: {self.totalNoContacts}")
    
    def update_contacts_attempted_till_now(self):
        self.contacts_attempted_till_now_label.config(text=f"Attempted: {self.totalContactsTillNow}  Successful: {self.successfulContacts}")

    def update_previous_message_to_contact_status(self):
        self.previous_message_to_contact_status_label.config(text=f"Sending message to {self.currentName} {self.messageStatus}.")
    
    def update_final_sent_status(self):
        self.total_sent_label.config(text=f"Total contacts attempted: {self.totalContactsTillNow}")
        self.successful_sent_label.config(text=f"Successfully sent to: {self.successfulContacts}")
        self.failedContacts = self.totalContactsTillNow - self.successfulContacts
        self.failed_sent_label.config(text=f"Failed sending to: {self.failedContacts}")

    def update_image_message(self, *args):
        if self.image_available_var.get() == "Yes":
            self.image_file_button.config(state=tk.NORMAL)
        else:
            self.imagePath = ""
            self.image_file_selected_label.config(text="")
            self.image_file_button.config(state=tk.DISABLED)

    def update_country_code_entry(self, *args):
        if self.are_contacts_with_country_code_var.get() == "No":
            self.country_code_entry.config(state=tk.NORMAL)
        else:
            self.country_code_entry.config(state=tk.DISABLED)

    def update_is_image_available_dropdown(self, *args):
        if self.browser_name_var.get() == "Firefox":
            self.image_available_combobox.config(state=tk.DISABLED)
            self.image_available_var.set("No")
        else:
            self.image_available_combobox.config(state=tk.NORMAL)

    def update_empty_input_label(self, text):
        self.empty_input_label.config(text=text, fg="red")
        self.empty_input_label.grid(row=8, column=0, columnspan=2, padx=10, pady=5)
        
    def submit_form(self):
        browserName = self.browser_name_var.get()
        
        
        isImageAvailable = self.display_to_code[self.image_available_var.get()]
        countryCode = self.country_code_entry.get()
        waitingTime = float(self.waiting_time_entry.get())  # Get the waiting time value
        if not browserName:
            self.update_empty_input_label("Please select browser from dropdown.")
        elif isImageAvailable == "Y" and not self.imagePath:
            self.update_empty_input_label("Please select image.")
        elif waitingTime <= 0:
            self.update_empty_input_label("Please enter waiting time > 0.")
        elif not self.contactsFilePath:
            self.update_empty_input_label("Please select contacts file.")
        elif not self.messageFilePath:
            self.update_empty_input_label("Please select message file.")
        else:
            # Forget all the earlier form elements
            self.browser_name_label.grid_forget()
            self.browser_name_combobox.grid_forget()

            self.image_available_label.grid_forget()
            self.image_available_combobox.grid_forget()

            self.image_name_label.grid_forget()
            self.image_file_button.pack_forget()
            self.image_file_selected_label.pack_forget()
            self.image_file_selection_frame.grid_forget()

            self.are_contacts_with_country_code_label.grid_forget()
            self.are_contacts_with_country_code_combobox.grid_forget()

            self.country_code_label.grid_forget()
            self.country_code_entry.grid_forget()

            self.waiting_time_label.grid_forget()
            self.waiting_time_entry.grid_forget()

            self.contact_label.grid_forget()
            self.contacts_file_button.pack_forget()
            self.contacts_file_selected_label.pack_forget()
            self.contacts_file_selection_frame.grid_forget()

            self.message_label.grid_forget()
            self.message_file_button.pack_forget()
            self.message_file_selected_label.pack_forget()
            self.message_file_selection_frame.grid_forget()

            self.empty_input_label.grid_forget()

            # self.image_message_label.grid_forget()
            self.submit_button.grid_forget()
            
            root.unbind("Return")
            
            # Create a label with the message
            self.continue_label = tk.Label(self.root, text="Press Continue when WhatsApp loads fully!")
            self.continue_label.pack(padx=20, pady=20)

            # Create the "Continue" button
            self.continue_btn = tk.Button(self.root, text="Continue", command=self.run_process)
            self.continue_btn.pack()
            self.root.bind("<Return>", lambda event: self.run_process())
        
            self.browserDriverMessageSender = BrowserDriverMessageSender(
                browserName=browserName, isImageAvailable=isImageAvailable, imagePath=self.imagePath, countryCode=countryCode, waitingTime=waitingTime, contactsFilePath=self.contactsFilePath, messageFilePath=self.messageFilePath
                )
            self.open_whatsapp_thread = threading.Thread(target=self.browserDriverMessageSender.open_whatsapp)
            self.open_whatsapp_thread.start()
            
    
    def run_process(self):
        self.continue_label.pack_forget()
        self.continue_btn.pack_forget()
        root.unbind("Return")
        # Create a label with the message
        
        self.total_no_contacts_label = tk.Label(self.root, text="")
        self.total_no_contacts_label.pack(padx=20, pady=(20, 10))

        self.contacts_attempted_till_now_label = tk.Label(self.root, text="Attempted: 0 Successful: 0") 
        self.contacts_attempted_till_now_label.pack(padx=20)
        
        self.progress = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
        self.progress.pack(pady=(0, 20))
        self.progress['maximum'] = 100

        self.previous_message_to_contact_status_label = tk.Label(self.root, text="Message Status: updating...", pady=10) #
        self.previous_message_to_contact_status_label.pack(padx=10)
        
        self.message_to_contact_label = tk.Label(self.root, text="") 
        self.message_to_contact_label.pack(padx=10, pady=(0,20))

        

        self.stop_label = tk.Label(self.root, text="Press Stop button to stop sending the message!")
        self.stop_label.pack(padx=20)

        # Create the "Stop" button
        self.stop_btn = tk.Button(self.root, text="Stop", command=self.stop_process)
        self.stop_btn.pack(pady=20)
        self.root.bind("<Return>", lambda event: self.stop_process())

        self.open_whatsapp_thread.join()
        
        is_send_whatsapp_message_thread_running.set()
        self.send_whatsapp_message_thread = threading.Thread(target=lambda: self.browserDriverMessageSender.start_send_message_process(self))
        self.send_whatsapp_message_thread.start()
    
    def stop_process(self):
        is_send_whatsapp_message_thread_running.clear()
        self.stop_btn.config(state=tk.DISABLED)
        

    def finish_process(self):
        self.stop_label.pack_forget()
        self.message_to_contact_label.pack_forget()
        self.previous_message_to_contact_status_label.pack_forget()
        self.progress.pack_forget()
        self.contacts_attempted_till_now_label.pack_forget()
        self.total_no_contacts_label.pack_forget()
        self.stop_btn.pack_forget()
        root.unbind("Return")

        self.message_upto_contact_label = tk.Label(self.root, text=f"Finished sending message upto contact {self.currentName}.")
        
        self.message_upto_contact_label.pack(pady=10)


        self.total_sent_label = tk.Label(self.root, text="") #
        self.total_sent_label.pack(pady=10)
        self.successful_sent_label = tk.Label(self.root, text="") #
        self.successful_sent_label.pack(pady=10)
        self.failed_sent_label = tk.Label(self.root, text="") #
        self.failed_sent_label.pack(pady=10)

        # Create a label with the message
        self.finish_label = tk.Label(self.root, text="Press Done to close the App.")
        self.finish_label.pack(padx=20, pady=10)

        self.root.bind("<Return>", lambda event: self.close_window())
        
        # Create the "Continue" button
        self.finish_btn = tk.Button(self.root, text="Done", command=self.close_window)
        self.finish_btn.pack(pady=20)
        self.update_final_sent_status()
        
        self.send_whatsapp_message_thread.join()
        self.browserDriverMessageSender.quit_driver()
        
        

class BrowserDriverMessageSender:
    def __init__(self, browserName, isImageAvailable, imagePath, countryCode, waitingTime, contactsFilePath, messageFilePath):
        self.isImageAvailable = isImageAvailable
        self.countryCode = countryCode
        self.waitingTime = waitingTime

        self.imageFilePath = imagePath
        
        self.contactsFileName   = os.path.basename(contactsFilePath)

        # take content of the files
        self.contactsList = []
        try:
            with open(contactsFilePath) as cf:
                self.contactsList.extend(csv.reader(cf))
        except FileNotFoundError as e:
            logging.error("File not found: %s", e)
        except csv.Error as e:
            logging.error("CSV Error while reading: %s", e)
        except Exception as e:
            logging.error("An error occurred while reading the file: %s", e)
        finally:
            try:
                cf.close()
            except:
                pass

        messagesList = ""
        try:
            with open(messageFilePath,"r", encoding="utf8") as mf:
                messagesList = mf.readlines()
        except FileNotFoundError as e:
            logging.error("File not found: %s", e)
        except Exception as e:
            logging.error("An error occurred while reading the file: %s", e)
        finally:
            try:
                mf.close()
            except:
                pass

        self.message = "%0A".join(messagesList)
        
        os_name = platform.system()

        # initialize browser
        if browserName == "Chrome":
            options = webdriver.ChromeOptions()
            if os_name == "Windows":
                cs = CS(executable_path="./chromedriver/chromedriver.exe")
                options.add_argument("--user-data-dir=C:/Users/"+getpass.getuser() + "/AppData/Local/Google/Chrome/User Data")
            elif os_name == "Darwin" or os_name == "macOS":
                cs = CS(executable_path="./chromedriver/darwinchromedriver")
                options.add_argument("--user-data-dir=/Users/"+getpass.getuser() + "/Library/Application Support/Google/Chrome/")
            options.add_argument("--profile-directory=Default")
            options.add_argument("--disable-extensions")
            options.add_argument('--start-maximized')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            self.driver = webdriver.Chrome(options=options, service=cs)
        
        elif browserName == "Firefox":
            fp = webdriver.FirefoxProfile(profile_directory='C:/Users/'+getpass.getuser() + '/AppData/Roaming/Mozilla/Firefox/Profiles/4ra3l7a3.default-release')
            options = webdriver.FirefoxOptions()
            options.profile = fp
            self.driver = webdriver.Firefox(options=options)

        elif browserName == "Edge":
            options = webdriver.EdgeOptions() 
            options.add_argument("--user-data-dir=C:/Users/"+getpass.getuser() + "/AppData/Local/Microsoft/Edge/User Data")
            options.add_argument("--profile-directory=Default")
            options.add_argument("--disable-extensions")
            options.add_argument('--start-maximized')
            self.driver = webdriver.Edge(options=options)

    
    def create_directory_if_not_exists(self, directory_path):
        path = Path(directory_path)
        if not path.exists():
            path.mkdir(parents=True, exist_ok=True)
            logging.info(f"Directory '{directory_path}' created.")
        else:
            logging.info(f"Directory '{directory_path}' already exists.")

    def open_whatsapp(self):
        linkWhatsAppCheck = 'https://web.whatsapp.com/'
        self.driver.get(linkWhatsAppCheck)
    
    def start_send_message_process(self, app):
        # Create Failed contact file in Failed folder
        directory_path = "./Failed"
        self.create_directory_if_not_exists(directory_path)
        try:
            fw = open('./Failed/' + 'Failed-'+self.contactsFileName, 'w+', newline='')
            writer = csv.writer(fw)
            # Your code for writing to the failed file goes here (if required)
        except Exception as e:
            logging.error("An error occurred while opening the failed file for writing: %s", e)

        messageStatusList = [""]
        totalNoContactsList = len(self.contactsList)
        app.totalNoContacts = totalNoContactsList
        app.update_total_no_contacts()

        for contact in self.contactsList:
            app.totalContactsTillNow += 1

            if is_send_whatsapp_message_thread_running.is_set():
                try:
                    app.currentName = contact[0]
                    app.update_message_to_contact()
                    phone = self.countryCode + contact[1]
                    link = 'https://web.whatsapp.com/send?phone=' + phone + '&text=' + 'Dear ' + contact[0]+'!  ' + '😀'+ '%0A' + self.message
                    self.driver.get(link)
                    self.driver.execute_script("window.onbeforeunload = function() {};")

                    time.sleep(self.waitingTime)
                    
                    if self.isImageAvailable in 'Y' or self.isImageAvailable in 'y':

                        documentButton = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div')
                        documentButton.click()
                        imageInputEle = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[5]/div/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[1]/li/div/input')
                        
                        imageInputEle.send_keys(self.imageFilePath)
                        time.sleep(2)
                        btn = self.driver.find_element(By.XPATH, '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div')
                    else :
                        btn = self.driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button')
                    
                    btn.click()
                    messageStatusList[0] = "succeeded"
                    app.successfulContacts +=1
                    logging.info("Success "+ contact[0] + '; ' + contact[1])
                except NoSuchElementException as e:
                    messageStatusList[0] = "failed"
                    logging.error(f"Failed {contact}.")
                    logging.error("Element not found. Check if the XPath is correct: %s", e)
                    try:
                        writer.writerow(contact)
                    except Exception as e:
                        logging.error(f"An error occurred while writing a row to the failed file:{str(e)}")
                except Exception as e: 
                    messageStatusList[0] = "failed"
                    logging.error(f"Failed {contact}. Error {str(e)}")
                    try:
                        writer.writerow(contact)
                    except Exception as e:
                        logging.error(f"An error occurred while writing a row to the failed file:{str(e)}")
                finally:
                    app.messageStatus = messageStatusList[0]
                    app.update_previous_message_to_contact_status()
                    app.progressBarValue = (app.totalContactsTillNow / totalNoContactsList) * 100
                    app.update_contacts_attempted_till_now()
                    app.update_progress_bar()
                    time.sleep(self.waitingTime/3)
            else:
                print("finished")
                break
        
        try:
            fw.close()
        except Exception as e:
            logging.error("An error occurred while closing the failed file: %s", e)
        app.root.after(0, app.finish_process)

    def quit_driver(self):
        self.driver.quit()


if __name__ == "__main__":
    # Set up logging
    logging.basicConfig(filename='error_log.txt', level=logging.ERROR,
                        format='%(asctime)s - %(levelname)s - %(message)s')

    logging.getLogger('selenium.webdriver').setLevel(logging.WARNING)
    is_send_whatsapp_message_thread_running = threading.Event()
    root = tk.Tk()
    app = App(root)
    # app.process_thread.start()  # Start the process in the background thread
    root.mainloop()