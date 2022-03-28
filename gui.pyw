import tkinter  as tk
from Auswertung import open_file 
import tkinter.filedialog as fd
import sys
import os

def restart_program():
    """Restarts the current program.
    Note: this function does not return. Any cleanup action (like
    saving data) must be done before calling this function."""
    form.destroy()
    os.startfile("gui.pyw")


        

def onclick():
    Status["text"] = "Fertig!"
    
def pfad(p):
    p.set(fd.askopenfilename())
    form.update()

def verzeichnis(v):
    v.set(fd.askdirectory())
    form.update()


form=tk.Tk()
form.wm_title("Auswertung Excel")
form.configure(background="White")

Head=tk.Label(form, text="Excel-Datei in viele Word-Templates einfügen")
Head.config( bg="White", font=("", 15))
Head.grid(row=0, column =0, columnspan = 3, sticky="N")

Restart= tk.Button(form, text="Restart", command=restart_program)
Restart.config( bg="RoyalBlue1", font=("", 15))
Restart.grid(row=0, column =3, columnspan = 3, sticky="N")

template=tk.StringVar()
Button_temp=tk.Button(form, text="Wählen Sie die Word-Vorlage-Datei:", bg="RoyalBlue1", command= lambda:pfad(template))
Button_temp.grid(row=1, column=1,  sticky="E", padx=5, pady=5, ipadx=5, ipady=5)

entry=tk.StringVar()
Button_entry=tk.Button(form, text="Wählen Sie die txt-Datei mit den Platzhalternamen:", bg="RoyalBlue1", command= lambda:pfad(entry))
Button_entry.grid(row=1, column=2,  sticky="E", padx=5, pady=5, ipadx=5, ipady=5)


way=tk.StringVar()
Button_in=tk.Button(form, text="Wählen Sie die Excel-Daten-Datei:", bg="RoyalBlue1", command= lambda:pfad(way))
Button_in.grid(row=1, column=3,  sticky="E", padx=5, pady=5, ipadx=5, ipady=5)

output=tk.StringVar()
Button_out=tk.Button(form, text="Wählen Sie den Ausgabe-Ordner:", bg="RoyalBlue1", command=lambda:verzeichnis(output))
Button_out.grid(row=1, column=4,  sticky="E", padx=5, pady=5, ipadx=5, ipady=5)

Button_go=tk.Button(form, text="Start (dauert etwa 5 Sekunden)", bg="LightSkyBlue1", command= lambda:[open_file(way.get(), output.get(), entry.get(), template.get()), onclick()])
Button_go.grid(row=2, column=1, sticky="E", padx=5, pady=5, ipadx=5, ipady=5)


Status=tk.Label(form, text="")
Status.config( bg="White", font=("", 15))
Status.grid(row=2, column =2, columnspan = 5, sticky="N")

form.mainloop()
