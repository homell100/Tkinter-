"""
Aspectes a solucionar:
1. Apareixen multiples finestres quan es grava l'arxiu Excel final i està obert
2. Hi ha un sleep.time(1) a "Input code" que hauria de mirar com millorar-l'ho
3. Assegurar-e que numeros grans com 1.000.000 son llegits correctament
4. Quan la columna dels codis no es diu Codi dona un error que no es interpretable, modificar-lo
per fer-lo mes comprensible.

Millores:
1. Casella on poder posar-li nom a l'arxiu Excel de resposta
2. Casella on es demani on guardar l'arxiu Excel de resposta
"""


import xlrd
import sys
import threading
import tkinter as tk
from tkinter import ttk, W, N, E, S, Button, filedialog, messagebox as msg, Text, END, Toplevel
from tkinter.ttk import Frame, Label, Combobox, Entry
from Data import months, years, rowErrors, width, height, cities
from Pandas import readExcel, write_on_excel
from API_Main import API_main

"""
#The present GUI has been created using Tkinter.
It mainly has three frames, structured vertically:
Top Frame contains: "Us element" button, "Municipi" button and range of dates buttons
Middle Frame contains: "Search for Excel file" button and "Stop process button"
Bottom Frame contains the display windows
"""

class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):

        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title("Descarrega dades de SIE-Observatori")
        self.parent.geometry(str(width) + "x" + str(height))

        ### Top Frame
        top_frame = Frame(self)
        top_frame.grid(row = 0, column = 0)
        # Left frame
        left_frame = Frame(top_frame, width = int(width/1.4), height = 100)
        left_frame.grid(row = 0, column = 0, padx=20, pady= 30)
        # Us element
        lbl_us_el = Label(left_frame, text="Us element")
        lbl_us_el.grid(row = 0, column = 0)
        self.combo_us_el = Combobox(left_frame, state="readonly")
        self.combo_us_el["values"] = ["Tots","Edifici", "Quadre"]
        self.combo_us_el.current(1)
        self.combo_us_el.grid(row = 1, column = 0)
        # Ciutats
        lbl_ciut = Label(left_frame, text="Ciutat")
        lbl_ciut.grid(row=2, column=0)
        self.combo_ciut = Combobox(left_frame, state="readonly")
        self.combo_ciut["values"] = cities
        self.combo_ciut.grid(row=3, column=0)
        # Right frame
        right_frame = Frame(top_frame, width =  int(width/2) , height = 100)
        right_frame.grid(row =0, column = 1, padx = 20, pady = 60, sticky = S)
        PeriodLabel = Label(right_frame, text="Periode de temps")
        PeriodLabel.grid(row=0, column=0, columnspan = 3)
        PeriodLabel.config(font=("Arial", 15))
        Label(right_frame, text="Inicial").grid(row=2, column = 0)
        Label(right_frame, text="Final").grid(row=3, column = 0)
        Label(right_frame, text="Mes").grid( row = 1, column = 1)
        Label(right_frame, text="Any").grid(row= 1, column = 2)
        # Initial month
        self.e1 = ttk.Combobox(right_frame, state="readonly")
        self.e1["values"] = months
        self.e1.current(0)
        self.e1.grid(row=2, column=1)
        # Initial year
        self.e2 = ttk.Combobox(right_frame, state="readonly")
        self.e2["values"] = years
        self.e2.current(1)
        self.e2.grid(row=2, column=2)
        # Final month
        self.e3 = ttk.Combobox(right_frame, state="readonly")
        self.e3["values"] = months
        self.e3.current(len(months)-1)
        self.e3.grid(row=3, column=1)
        # Final year
        self.e4 = ttk.Combobox(right_frame, state="readonly")
        self.e4["values"] = years
        self.e4.current(1)
        self.e4.grid(row=3, column=2)

        ## Middle frame
        middle_frame = Frame(self)
        middle_frame.grid(row=1, column = 0)
        # Searching excel Button
        search_btn = Button(middle_frame, text = "Buscar arxiu Excel amb CODIS SIE", command = self.start_searching, bg = "lightblue")
        search_btn.grid(row = 0, column = 0, pady = 5, padx = 10)
        # Stop button
        stop_btn = Button(middle_frame, text="Parar procés", command=self.stop_searching, bg="red")
        stop_btn.grid(row=0, column=1, pady = 10, padx = 10)

        ## Bottom frame
        bottom_frame = Frame(self)
        bottom_frame.grid(row=2, column=0)
        # Log Screen
        self.log = Text(bottom_frame, width =  80 , height = 15, padx = 10, pady = 10 )
        self.log.grid(row = 1, column = 0)

        class PrintToLog(object):
            def __init__(self, log):
                self.log = log
            def write(self,s):
                self.log.insert(END, s)
                self.log.see(END)
            def flush(self):
                pass

        sys.stdout = PrintToLog(self.log)

    def start_searching(self):

        self.log.delete('1.0', END)
        self.stop_event = threading.Event()
        class Config(object):
            def __init__(self, initial_date, final_date, us_element, ciutat):
                self.initial_date = initial_date
                self.final_date = final_date
                self.us_element = us_element
                self.ciutat = ciutat

        initial_date = [self.e1.get(), self.e2.get()]
        final_date = [self.e3.get(), self.e4.get()]
        us_element = self.combo_us_el.get()
        ciutat = self.combo_ciut.get()

        if ciutat != "":
            self.config = Config(initial_date, final_date, us_element, ciutat)
            thread = threading.Thread(target=self.getExcelAndExecuteApi)
            thread.start()
        else:
            msg.showinfo(" ","No hi ha cap ciutat escollida")

    """
    Class that controls the stop of the process
    """
    def stop_searching(self):
        msg.showinfo("", "El procés ha sigut aturat. Els càlculs es pararan a partir del següent codi...")
        self.stop_event.set()
    """
    Main class that reads the input excel, scraps the web and saves the results in another excel file
    """
    def getExcelAndExecuteApi(self):
        import_file_path = filedialog.askopenfilename()

        try:
            df = readExcel(import_file_path)
            try:
                result = API_main(df, self.config, self.stop_event)
                while True:
                    try:
                        write_on_excel(result)
                        print("Excel amb els resultats escrit correctament.")
                        break
                    except PermissionError:
                        self.displayPermissionError()

                if len(rowErrors) < 1:
                    msg.showinfo(" ", "Procés completat sense errors")
                else:
                    print("*" * 20, "\n", "Codis on s'han trobat errors :", rowErrors, "\n", "*" * 20)
                    msg.showinfo(" ", "Procés completat amb errors")
            except:
                print("S'ha trobat un error: \n", sys.exc_info()[0])
                print("Codis on s'han trobat errors :", rowErrors)

                msg.showinfo("", "Proces aturat per errors")

        except xlrd.biffh.XLRDError:
            msg.showinfo("","La fulla del Excel ha llegir s'ha d'anomenar API")

    """
    Need finilisation of code
    """
    def displayPermissionError(self):
        newWin = Toplevel(root)
        var = tk.IntVar()
        # Text
        display = Label(newWin, text = "El archivo Anwser.xlsx està obert, siusplau tanca'l i prem 'Torna-ho  a intentar'"  )
        display.grid(row = 0, column = 0)
        button1 = Button(newWin, text = "Tornar-ho a intentar", command = lambda: var.set(1))
        button1.grid(row = 1, column = 0)
        button1.wait_variable(var)
        # Button1
        # Button2
        pass

if __name__ == "__main__":
    root = tk.Tk()
    MainApplication(root). pack(side="top", fill="both", expand=True)
    root.mainloop()

