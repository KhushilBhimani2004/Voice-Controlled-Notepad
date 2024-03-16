from tkinter import *
from tkinter.messagebox import showinfo
from tkinter.filedialog import askopenfilename, asksaveasfilename, asksaveasfile
from tkinter import messagebox
from tkinter import colorchooser as cc
from tkinter import font
from tkinter.font import Font
from datetime import datetime
import os
import win32api
import speech_recognition as sr
import pyttsx3
import wikipedia


engine = pyttsx3.init()
engine.setProperty('rate', 120)

# functions
def newFile(x=''):
    global file
    root.title("Untitled - Notepad")
    file = None
    TextArea.delete(1.0, END) #delete all till 1st line and 0th character ==> 1.0

def openFile(x=''):
    global file
    file = askopenfilename(defaultextension='.txt', filetypes=[("All Files", "*.*"), ("Text Documents", "*.txt"), ("Python Files", "*.py")])
    if file == "":
        file = None
    else:
        root.title(os.path.basename(file)+" - Notepad")
        TextArea.delete(1.0, END)
        f = open(file, "r")
        TextArea.insert(1.0, f.read())
        f.close()
def saveFile(x=''):
    global file
    if file == None:
        file = asksaveasfilename(initialfile='Untitled.txt', defaultextension='.txt', filetypes=[("All Files", "*.*"), ("Text Documents", "*.txt"), ("Python Files", "*.py")])
        if file=="":
            file = None
        else:
            f = open(file, "w")
            f.write(TextArea.get(1.0, END))
            f.close()
            root.title(os.path.basename(file)+" - Notepad")
            print("file saved")

    else:
        f = open(file, "w")
        f.write(TextArea.get(1.0, END))
        f.close()

def saveasFile(x=''):
    global fd
    fd = asksaveasfile(mode = 'w', defaultextension='.txt', filetypes=[("Text Documents", "*.*"), ("Text Documents", "*.txt"), ("Python Files", "*.py")])
    if fd!= None:
        data = TextArea.get('1.0', END)
    try:
        fd.write(data)
    except:
        messagebox.showerror(title="Error", message = "Not able to save file!")

def closeApp(x=''):
    root.destroy()
def cut():
    TextArea.event_generate(("<<Cut>>"))
def copy():
    TextArea.event_generate(("<<Copy>>"))
def paste():
    TextArea.event_generate(("<<Paste>>"))
def about():
    showinfo("Voice Controlled Notepad","Our Notepad makes your work more easy and time savy by using the advance features of voice control!")
def printFile(x=''):
    global file
    file = askopenfilename(initialfile='Untitled.txt', defaultextension='.txt', filetypes=[("All Files", "*.*"), ("Text Documents", "*.txt"), ("Python Files", "*.py")])
    
    if file:
        # Print Hard Copy of File
        win32api.ShellExecute(0, "print", file, None, ".", 0)
def fbold(x=''):
    bold_font=font.Font(TextArea, TextArea.cget("font"))
    bold_font.configure(weight="bold")

    TextArea.tag_configure("bold", font=bold_font)
    current_tags = TextArea.tag_names("sel.first")
    if "bold" in current_tags:
        TextArea.tag_remove("bold", "sel.first", "sel.last")
    else:
        TextArea.tag_add("bold","sel.first", "sel.last")

def fitalic(x=''):
    italics_font=font.Font(TextArea, TextArea.cget("font"))
    italics_font.configure(slant="italic")

    TextArea.tag_configure("italic", font=italics_font)
    current_tags = TextArea.tag_names("sel.first")
    if "italic" in current_tags:
        TextArea.tag_remove("italic", "sel.first", "sel.last")
    else:
        TextArea.tag_add("italic","sel.first", "sel.last")
def dateandtime():
    TextArea.insert(END, datetime.now())
def selectall(x=''):
    TextArea.tag_add('sel', '1.0', 'end')
def t_col(x=''):
    my_color = cc.askcolor()[1]
    if my_color:
        TextArea.config(fg=my_color)
def bg_col(x=''):
    my_color = cc.askcolor()[1]
    if my_color:
        TextArea.config(bg=my_color)


def vc():
    showinfo("Voice Controlled Notepad"," Please press the button at the bottom or press 'Alt+V' Buttons to enable voice control.")

if __name__ == '__main__':
    root = Tk()
    root.title("Untitled-Notepad")
    root.wm_iconbitmap("notepad.png")
    root.geometry("600x400")


    #textarea
    TextArea = Text(root, font='TimesNewRoman 13')
    file = None
    TextArea.pack(fill=BOTH, expand=True)
    #speech recognition starts 
    def transcribe_audio(x=''):

        recognizer = sr.Recognizer()
        print('listening')
        engine.say("I am listening, give your command")
        engine.runAndWait()
        with sr.Microphone() as source:
            audio = recognizer.listen(source)
            try:
                text = recognizer.recognize_google(audio)
                TextArea.insert(END, text + "\n")

                if 'close app' in text.lower():
                    closeApp()
                if 'open file' in text.lower():
                    openFile()
                if 'sam save file' in text.lower():
                    saveFile()
                if 'sam save file as' in text.lower():
                    saveasFile()
                if 'sam open new file' in text.lower():
                    newFile()
                if 'info' in text.lower():
                    txt = text[4::]
                    try:
                        page = wikipedia.page(txt)
                        TextArea.insert(END,page.title)
                        TextArea.insert(END,page.summary)
                    except wikipedia.exceptions.PageError:
                        engine.say("Sorry, couldn't find any information about " + txt)
                        engine.runAndWait()
                if 'google' in text.lower():
                    txt = text[7::]
                    url = f"https://www.google.com/search?q={txt}"
                    os.system(f"start {url}")
                    engine.say("Googling it")
                    engine.runAndWait()
                    # TextArea.insert(END, info)
                if 'x' in text.lower():
                    # print('entered')
                    # test
                    global file
                    txt = text[:-1:1] #my demo file x
                    fname = txt+'.txt'
                    root.title(os.path.basename(fname)+" - Notepad")
                    TextArea.delete(1.0, END)
                    f = open(fname, "r")
                    TextArea.insert(1.0, f.read())
                    f.close()
                    # end
                if 'save' in text.lower():
                    # test
                    global file
                    txt = text[:-5:1]
                    if len(txt) < 1:
                        engine.say("Please enter a valid name!")
                        engine.runAndWait()
                        # print('')
                    else:
                        fname = txt+'.txt'
                        f = open(fname, "w")
                        f.write(TextArea.get(1.0, END))
                        f.close()
                        root.title(os.path.basename(fname)+" - Notepad")
                        engine.say(" file saved")
                        engine.runAndWait()
                        # print("")
                    # end
                # else:
                # pass
            except sr.UnknownValueError:
                engine.say(" Sorry, could not understand audio.")
                engine.runAndWait()
                # TextArea.insert(END, "Sorry, could not understand audio.\n")
            except sr.RequestError as e:
                TextArea.insert(END, "Could not request results from Google Speech Recognition service: {0}.\n".format(e))
    
    button = Button(root, text="Voice Control", command=transcribe_audio)
    button.pack()
    #ends


    # adding scrollbar
    ScrollBar = Scrollbar(TextArea)
    ScrollBar.pack(side=RIGHT, fill=Y)  
    ScrollBar.config(command=TextArea.yview)
    TextArea.config(yscrollcommand=ScrollBar.set, undo=True)

    # menubar
    MenuBar = Menu(root)
    FileMenu = Menu(MenuBar, tearoff=0)

    # to open new file
    FileMenu.add_command(label="New", command=newFile)
    # to open existing file
    FileMenu.add_command(label="Open",command=openFile)

    # to save the current file
    FileMenu.add_command(label="Save", command=saveFile)
    FileMenu.add_command(label="Save as...", command=saveasFile)
    FileMenu.add_separator()
    FileMenu.add_cascade(label="Print", command=printFile)
    FileMenu.add_separator()
    FileMenu.add_command(label="Exit", command=closeApp)
    MenuBar.add_cascade(label = "File", menu=FileMenu)
    root.config(menu=MenuBar)

    #edit menu
    EditMenu= Menu(MenuBar, tearoff=0)
    # undo redo
    EditMenu.add_cascade(label="Undo", command=TextArea.edit_undo)
    EditMenu.add_cascade(label="Redo", command=TextArea.edit_redo)
    EditMenu.add_separator()
    #copy paste cut
    EditMenu.add_command(label="Copy", command=copy)
    EditMenu.add_command(label="Paste", command=paste)
    EditMenu.add_command(label="Cut", command=cut)
    #font
    EditMenu.add_command(label="Bold", command=fbold)
    EditMenu.add_command(label="Itallic", command=fitalic)
    # EditMenu.add_command(label="Underlined", command=ffam)
    EditMenu.add_separator()
    EditMenu.add_command(label="Date/Time", command=dateandtime)
    EditMenu.add_command(label="Select All", command=selectall)
    
    MenuBar.add_cascade(label="Edit",menu=EditMenu)
    #ends
    # color menu
    ColorMenu = Menu(MenuBar, tearoff=0)
    ColorMenu.add_command(label="Text", command=t_col)
    ColorMenu.add_command(label="Background", command=bg_col)
    MenuBar.add_cascade(label="Color", menu=ColorMenu)

    # ends
    # voice control menu
    VoiceMenu = Menu(MenuBar, tearoff=0)
    VoiceMenu.add_command(label='Voice Control', command = vc)
    MenuBar.add_cascade(label='Voice Access', menu = VoiceMenu)
    # help menu
    HelpMenu = Menu(MenuBar, tearoff=0)
    # about me
    HelpMenu.add_command(label='About', command=about)
    MenuBar.add_cascade(label="Help", menu=HelpMenu)
    # ends

    # shortcut keys
    root.bind('<Alt-x>', closeApp)
    root.bind('<Control-p>', printFile)
    root.bind('<Control-Shift-s>', saveasFile)
    root.bind('<Control-o>', openFile)
    root.bind('<Control-s>', saveFile)
    root.bind('<Alt-a>', about)
    root.bind('<Control-n>', newFile)
    root.bind('<Control-b>', fbold)    
    root.bind('<Alt-i>', fitalic)
    root.bind('<Alt-a>', selectall)
    root.bind('<Alt-c>',t_col)
    root.bind('<Alt-b>',bg_col)
    root.bind('<Alt-v>', transcribe_audio)
    engine.say("Welcome to Voice Control Notepad, please click on the button or press Alt+V button to activate Voice Control.")
    engine.runAndWait()
    root.mainloop()
    engine.say("Thanks for using our system!")
    engine.runAndWait()