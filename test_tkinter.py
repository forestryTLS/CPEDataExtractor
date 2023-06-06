from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import tkinter as tk

class BrowserController:
    def __init__(self):
        self.driver = None

    def start_browser(self):
        self.driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
        self.driver.get('http://www.google.com')

    def print_title(self):
        if self.driver:
            title = self.driver.title
            print(f'Page title is: {title}')

    def quit(self):
        if self.driver:
            self.driver.quit()

if __name__ == '__main__':
    root = tk.Tk()

    controller = BrowserController()

    start_button = tk.Button(root, text='Start browser', command=controller.start_browser)
    start_button.pack()

    title_button = tk.Button(root, text='Print page title', command=controller.print_title)
    title_button.pack()

    # run the GUI
    root.mainloop()

    controller.quit()
