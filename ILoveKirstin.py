from tkinter import *
import os
import tkinter.filedialog as fd
import ILKDivide as ilkd
import ILKFormat as ilkf
import ILKWebcheck as ilkw

window = Tk()
my_image = PhotoImage(file="ilk-app-logo.png")   #logo image
my_icon = PhotoImage(file='alt-ilk-app-icon.png')#icon image
window.iconphoto(False, my_icon)                 #icon image 

def close_window():
    window.destroy()
    exit
    pass

def browse_file():
    currdir = os.getcwd()
    tempdir = fd.askopenfilename(parent=window, initialdir=currdir, title='Please choose your .xlsx file')
    if len(tempdir) > 0:
        return(tempdir)

def norm_path(path):
    norm_p = os.path.basename(os.path.normpath(str(path)))
    return norm_p

def leave_home():
    for El in home_El:
        El.place_forget()

def return_home(return_list):
    for i in return_list:
        i.place_forget()
    quit_b.place(x = 260, y = 10)
    info_b.place(x = 4, y = 10)
    format_b.place(x = 4, y = 160)
    webcheck_b.place(x = 60, y = 160)
    divide_b.place(x = 140, y = 160)
    app_img.place(x=0, y=50)

## Module Functions ##

def open_info():
    window2 = Toplevel(height=615, width=750)
    info_paragraph = Label(window2, justify= LEFT, wraplength= "750", text="Introduction: \n \n This program (titled \"I Love Kirstin\") was designed to be a tool for my wonderful wife, Kirstin. She is the love of my life, my rock, and the one person whom I rely on most of all. I hope this program will help her spend less time doing the boring parts of her job. \n \n ILK Modules Explained: \n \n The FORMAT button takes an .xlsx file and changes the columns in the file to fit the user's needs. \n The WEBCHECK button allows the user to create an automated script that takes the values from the .xlsx and checks them against the same values on a given website. Webcheck is by far the most complicated part of the project and has a small variety of extra tools designed to help the end-user create working automation scripts. The key philosophy of Webcheck is to \"give non-technical users the tools to teach their computers how to do their jobs.\" In my opinion, the human element will never be replaceable by computers because critical thinking is essential. Instead, ILK makes its users more efficient in their day-to-day jobs. \n The DIVIDE button seperates a .xlsx file row by row into any number of new .xlsx files. The user can choose a column to sort the rows by before dividing the document. \n \n Error Trapping: \n \n It\'s important to know that the current version of ILK is NOT ACTIVELY OPTIMIZED FOR ERROR TRAPPING. This means that if you are using the tool incorrectly, it won't work for you as intended. In fact, you may not even realize when the tool is working incorrectly. Thus, this technology is defined as \"USE AT YOUR OWN RISK\" and the developers / contributors make no claim to the accuracy of the tool itself or the files that are created. \n \n Dependencies: \n \n This project relies on Python, Selenium, OpenPyxl, and Google Chrome 95 to function properly. Other modules are used, but these are the main ones. In the future, a VPN module may be added as well. \n \n ILK Key Concepts: \n \n In most cases, the user will be initially prompted to select a HEADER ROW. This tells ILK what row to skip to start the specific process(format, webcheck or divide). In addition, ILK will never make destructive changes to the original excel files. Instead, it will create new files (prefixed with FMT, DIV, WEB, SCHEME and ERROR) in the parent folder where the original .xlsx file was selected. Lastly, ILK Webcheck was designed around the CSS Selector locator exclusively because of its broad applicability to many websites. However, in some situations, Webcheck will be unable to automate/scan a website or app. \n \n Licensing: \n \n ILK is offered as free software under the MIT License. Selenium is offered under the Apache 2.0 license. No proprietary information/business process is included in ILK.")
    info_paragraph.place(x = 4, y = 10)
    home_b = Button(window2, text = "Hide Info", command = lambda: window2.destroy(), bg = "gray", fg = "white")
    home_b.place(x = 685, y = 10)
    #info_window.iconphoto(False, my_icon)
    #info_window.title('I Love Kirstin')
    #info_window.geometry("300x300+20+10")
    #info_window.mainloop()

def open_format(tempdir="none"): #work on pulling var from entry
    leave_home()
    if tempdir == "none":
        tempdir = browse_file()
    else:
        pass
    info_paragraph = Label(justify=LEFT, text="You Chose This File:" + "\n" + "\n" + norm_path(tempdir))
    info_paragraph.place(x = 4, y = 20)
    #
    info_paragraph1 = Label(justify=LEFT, text="Columns to Hide:")
    info_paragraph1.place(x = 4, y = 100)
    #
    format_e = Entry(window)
    format_e.place(x = 10, y = 125)
    format_e.insert(0, "A B C")
    #
    info_paragraph2 = Label(justify=LEFT, text="Columns to Widen:")
    info_paragraph2.place(x = 144, y = 100)
    #
    format_e2 = Entry(window)
    format_e2.place(x = 150, y = 125)
    format_e2.insert(0, "D E F")
    #
    info_paragraph3 = Label(justify=LEFT, text="Columns to Unformat:")
    info_paragraph3.place(x = 4, y = 150)
    #
    format_e3 = Entry(window)
    format_e3.place(x = 10, y = 175)
    format_e3.insert(0, "")
    #
    info_paragraph4 = Label(justify=LEFT, text="Unformat All? YES / NO")
    info_paragraph4.place(x = 144, y = 150)
    #
    format_e4 = Entry(window)
    format_e4.place(x = 150, y = 175)
    format_e4.insert(0, "NO")
    #
    info_paragraph5 = Label(justify=LEFT, text="Base Name:")
    info_paragraph5.place(x = 4, y = 200)
    #
    format_e5 = Entry(window)
    format_e5.place(x = 10, y = 225)
    format_e5.insert(0, "")
    #
    format_b = Button(text = "Submit", command = lambda: ilkf.main_format_function(tempdir, format_e.get(), format_e2.get(), format_e3.get(), format_e4.get(), format_e5.get()), bg = "green", fg = "white")
    format_b.place(x = 4, y = 250)
    #
    home_b = Button(text = "Home", command = lambda: return_home([home_b, format_b, format_e, format_e2, format_e3, format_e4, format_e5, info_paragraph, info_paragraph1, info_paragraph2, info_paragraph3, info_paragraph4, info_paragraph5]), bg = "gray", fg = "white")
    home_b.place(x = 240, y = 10)
    #hide_list, grow_list, clear_list, clear_all,

def open_webcheck(tempdir="none"):
    leave_home()
    if tempdir == "none":
        tempdir = browse_file()
    else:
        pass
    info_paragraph = Label(justify=LEFT, text="You Chose This File:" + "\n" + "\n" + norm_path(tempdir) + "\n" + "\n" + "Search Site:")
    info_paragraph.place(x = 4, y = 20)
    #
    webcheck_e = Entry(window, width=25)
    webcheck_e.place(x = 10, y = 105)
    webcheck_e.insert(0, "https://svc.mt.gov/dor/property/prc") #https://www.google.com/
    #
    info_paragraph2 = Label(justify=LEFT, text="Header Row:")
    info_paragraph2.place(x = 4, y = 125)
    #
    webcheck_e2 = Entry(window)
    webcheck_e2.place(x = 10, y = 150)
    webcheck_e2.insert(0, "10")
    #
    info_paragraph3 = Label(justify=LEFT, text="Search Column:")
    info_paragraph3.place(x = 4, y = 175)
    #
    webcheck_e3 = Entry(window)
    webcheck_e3.place(x = 10, y = 200)
    webcheck_e3.insert(0, "F")
    #
    info_paragraph4 = Label(justify=LEFT, text="Base Name:")
    info_paragraph4.place(x = 144, y = 175)
    #
    webcheck_e4 = Entry(window)
    webcheck_e4.place(x = 150, y = 200)
    webcheck_e4.insert(0, "New Sheet")
    #
    #main_webcheck_function(file_p, base_webpage, int(header_row), search_column, given_name="default")
    webcheck_b = Button(text = "Submit", command = lambda: ilkw.main_webcheck_function(tempdir, webcheck_e.get(), int(webcheck_e2.get()), webcheck_e3.get(), webcheck_e4.get(), window), bg = "green", fg = "white")
    webcheck_b.place(x = 4, y = 250)
    #
    home_b = Button(text = "Home", command = lambda: return_home([home_b, info_paragraph, info_paragraph2, info_paragraph3, info_paragraph4, webcheck_e, webcheck_e2, webcheck_e3, webcheck_e4, webcheck_b]), bg = "gray", fg = "white")
    home_b.place(x = 240, y = 10)

def open_divide(tempdir="none"): #work on pulling var from entry
    leave_home()
    if tempdir == "none":
        tempdir = browse_file()
    else:
        pass
    info_paragraph = Label(justify=LEFT, text="You Chose This File:" + "\n" + "\n" + norm_path(tempdir) + "\n" + "\n" + "\n" + "# of Spreadsheets:" + "                   Header Row:")
    info_paragraph.place(x = 4, y = 20)
    #
    #number of files
    divide_e = Entry(window)
    divide_e.place(x = 10, y = 125)
    #
    #header row
    divide_e3 = Entry(window)
    divide_e3.place(x = 160, y = 125)
    divide_e3.insert(0, 10)
    #
    info_paragraph2 = Label(justify=LEFT, text="Base Name:" + "                               Sort Column:")
    info_paragraph2.place(x = 4, y = 150)
    #
    #Base name
    divide_e2 = Entry(window)
    divide_e2.place(x = 10, y = 175)
    #
    #Sort column
    divide_e4 = Entry(window)
    divide_e4.place(x = 160, y = 175)
    divide_e4.insert(0, "AN")
    #
    divide_b = Button(text = "Submit", command = lambda: ilkd.main_divide_function(tempdir, divide_e.get(), divide_e2.get(), divide_e3.get(), divide_e4.get() ), bg = "green", fg = "white")
    divide_b.place(x = 4, y = 250)
    #
    home_b = Button(text = "Home", command = lambda: return_home([home_b, info_paragraph, divide_b, divide_e, divide_e2, divide_e3, divide_e4, info_paragraph2]), bg = "gray", fg = "white")
    home_b.place(x = 240, y = 10)

def open_scan(open_module):
    if open_module == "format":
        print("format")
        open_format("none")
    elif open_module == "webcheck":
        print("webcheck")
        open_webcheck("none")
    elif open_module == "divide":
        print("divide")
        open_divide("none")
    else:
        print("error")
    
#Home Screen Widgets:
quit_b = Button(text = "Quit", command = lambda: close_window(), bg = "red", fg = "white")
quit_b.place(x = 260, y = 10)

info_b = Button(text = "Info", command = lambda: open_info(), bg = "gray", fg = "white")
info_b.place(x = 4, y = 10)

format_b = Button(text = "Format", command = lambda: open_scan("format"), bg = "#585587", fg = "white")
format_b.place(x = 4, y = 160)

webcheck_b = Button(text = "Web Check", command = lambda: open_scan("webcheck"), bg = "#6d96be", fg = "white")
webcheck_b.place(x = 60, y = 160)

divide_b = Button(text = "Divide", command = lambda: open_scan("divide"), bg = "#986466", fg = "white")
divide_b.place(x = 140, y = 160)

app_img = Label(image=my_image) #logo image
app_img.place(x=0, y=50)        #logo image

home_El = [quit_b, info_b, format_b, webcheck_b, divide_b, app_img]

window.title('I Love Kirstin')
window.geometry("300x300+20+10")
window.mainloop()