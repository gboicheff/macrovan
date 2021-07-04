import tkinter as tk
import tkinter.font as tkFont
from utils import *
from printing_steps import *
import time

def printNowButton():
    print('print button clicked')
    # Start from 'List' screen
    # Get list name and then clear the field
    list_name = list_name_entry.get()
    if (list_name == ""):
        no_name_warning_text.pack()
    else:
        if (no_name_warning_text.winfo_ismapped()):
            no_name_warning_text.pack_forget()
        list_name_entry.delete(0, tk.END)
        script_name = "*2020 D68 Aug Primary"
        print_controller(driver, list_name, script_name)
        continue_instructions.pack()
        continue_button.pack()


def continueButton():
    print('continue button clicked')
    final_selections_submit(driver)
    continue_instructions.pack_forget()
    continue_button.pack_forget()

# def exitButton():
#     print('exit button clicked')

    
# Main function for testing

if __name__ == '__main__':

    #Suppress all print statements
    #disable_print()

    window = tk.Tk()
    window.title('GUIvan')


    # Create driver and login
    driver = start_driver()
    get_page(driver)
    driver.implicitly_wait(10)
    # take out login for release
    login_to_page(driver)


    #Override the exit button of the tk window.
    #window.wm_protocol("WM_DELETE_WINDOW", lambda: exit_program(window, driver))

    # Create GUI
    #Create custom font to adjust size
    custom_font = tkFont.Font(family="Courier", size=17, weight="bold")
    # Create Labels for instructions
    instructions1 = tk.Label(
        text="Please enter a name for this print file (pdf)",
        font=custom_font,
        padx=10,
        pady=5
    )
    instructions2 = tk.Label(
        text="To automatically click VAN print icon \nclick 'Print Now' button",
        font=custom_font,
        pady=5
    )
    continue_instructions = tk.Label(
        text="Please make sure everything is correct. After reviewing press 'Continue'",
        font=custom_font,
        pady=5
    )
    no_name_warning_text = tk.Label(
        text="Please enter a name for the print file",
        fg = "#ff0000",
        font=custom_font,
        pady=5
    )
    # exit_instructions = tk.Label(
    #     text="When you finish, click 'Exit'."
    # )

    # Create text Input
    list_name_entry = tk.Entry(
        width=50
    )

    # Create Buttons
    print_now_button = tk.Button(
        text="Print Now",
        width=25,
        height=5,
        fg="#ffffff",
        bg="#419cd1",
        command=printNowButton,
        font=custom_font
    )
    continue_button = tk.Button(
        text="Continue",
        width=25,
        height=5,
        fg="#ffffff",
        bg="#419cd1",
        command=continueButton,
        font=custom_font
    )
    # exit_button = tk.Button(
    #     text="Exit",
    #     width=15,
    #     height=5,
    #     fg="snow",
    #     bg="steel blue",
    #     command=exitButton
    # )

    # pack everything into GUI
    instructions1.pack()
    no_name_warning_text.pack()
    no_name_warning_text.pack_forget()
    list_name_entry.pack()
    instructions2.pack()
    print_now_button.pack()
    # continue_instructions.pack()
    # continue_button.pack()
    # exit_instructions.pack()
    # exit_button.pack()
    #window.after(1500,lambda: check_browser(window, driver))
    window.mainloop()
