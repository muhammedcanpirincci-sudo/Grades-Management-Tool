from Tkinter import *  #Importing modules
import ttk,xlrd,tkFileDialog     #Importing modules
from xlwt import * #Importing modules



class Student(): #Defining first class
    def __init__(self,ID,Name,Section,Dept,GPA,MP1,MP2,MP3,MT,FINAL): #Defining attributes of student
        self.Name = Name
        self.Section = Section
        self.ID = ID
        self.Dept = Dept
        self.GPA = GPA
        self.MP1 = MP1
        self.MP2 = MP2
        self.MP3 = MP3
        self.MT = MT
        self.FINAL = FINAL



class StudentList():#defining second class
    def __init__(self,Students): #defining attribute that is the dictionary of student objects
        self.Students = Students
    def reading_excel(self):#function to read the excel file imported and adding to the dictionary
        self.Xls_File_Path = tkFileDialog.askopenfilename(initialdir="/", title="Select file",
                                                          filetypes=(("xlsx files", ".xlsx"), ("all files", ".*")))#asks user for file
        self.workbook = xlrd.open_workbook(self.Xls_File_Path)#opens the file the user has selected

        self.sheet = self.workbook.sheet_by_index(0)


        for x in range(len(self.sheet.col_values(0)) - 1):#---------------------------------------------scanning columns and rows and reading data----------------------------------------------
            self.list = [] #list that will later hold values of each row and be used to create student object
            for i in self.sheet.row_values(x + 1):
                self.list.append(i)

            dict_of_grades = {5:self.list[5],6:self.list[6],7:self.list[7],8:self.list[8],9:self.list[9]} #DICT OF MP1 MP2 MP3 MT FINAL GRADES
            for item in dict_of_grades:    #iterates through each key in dict

                #code that checks if grades of mp1,2,3,mt,final has numbers and if they do it converts them to int
                if type(dict_of_grades[item])== float:
                    self.list[item] = int(self.list[item])

            # Inserting into dictionary of student objects
            self.Students[str(self.list[0]).split(".")[0]] \
                    = Student(self.list[0], self.list[1].split(" "), self.list[2], self.list[3], self.list[4], self.list[5],
                        self.list[6], self.list[7], self.list[8], self.list[9])






class GUI(Frame):       #defining third class
    def __init__(self, parent):         #defining attributes of third class
        self.parent = parent
        self.StudentList = StudentList({})          #creating object of second class to access dictionary
        self.control_variable = 0           #variable that checks if file has been imported
        Frame.__init__(self, parent)
        self.initUI(parent)             #calling the ui function which displays everything


    def initUI(self,parent):     #creating a function that has all widgets
        self.first_grade_entry = StringVar()   #variables that hold the value of certain widgets
        self.second_grade_entry = StringVar()  #
        self.third_grade_entry = StringVar()   #
        self.forth_grade_entry = StringVar()   #
        self.fifth_grade_entry = StringVar()   #
        self.Checkbox = StringVar()            #
        self.Checkbox.set("e")                 # Setting Default value for checkbox
        self.var_of_filename = StringVar()     #

        self.label_1 = Label(self, text="Grades Management Tool", bg="#36d16c", fg="white", font=("Helvetica", 18,"bold"),width=46)# ============CREATING AND PLACING WIDGETS=========
        self.label_1.grid(sticky=W, row=0, column=0)#MAIN TITLE LABEL

        self.Selectfile_button = Button(self, text="Select File",font = ("Helvetica",10,"bold"),command = self.file_selection)
        self.Selectfile_button.grid(sticky=W, row=2, column=0, pady=8, padx=85)#SELECT FILE BUTTON

        self.Treeview = ttk.Treeview(self,height = 10)#TreeView widget
        self.Treeview['show'] = 'headings'
        self.Treeview['columns'] = ("ID", "Name", "Surname")#---------setting default titles for each column ---------
        self.Treeview.heading('ID', text='ID')
        self.Treeview.heading('Name', text='Name')
        self.Treeview.heading('Surname', text='Surname')
        self.Treeview.column('ID', width=90)
        self.Treeview.column('Name', width=90)
        self.Treeview.column('Surname', width=90)


        self.Treeview.grid(sticky=W, row=4, column=0, pady=2, padx=10)

        self.Show_Data_Butoon = Button(self, text="Show Data",font = ("Helvetica",8,"bold"),command = self.show_data)#SHOW DATA BUTTON
        self.Show_Data_Butoon.grid(sticky="NW", row=4, column=0,pady = 70, padx=290)

        self.newframe1 = Frame(self, borderwidth=2, relief=GROOVE)#Frame for student details

        self.newframe1.grid(row=4, column=0,ipady = 10,padx=365)
                                                                                                #-------------STDUENT DETAILS TAB------------
        self.Baslik_For_Newframe = Label(self.newframe1, text="Student Details:", bg="#2c8973", fg="white",
                                         font=("Helvetica", 9,"bold"), height=1, width=44)
        self.Baslik_For_Newframe.grid(sticky=W, row=0, column=0)

        self.name = Label(self.newframe1, text="Name",font = ("Helvetica",9,"bold"))
        self.name.grid(sticky=W, row=1, column=0)

        self.name_entry = Label(self.newframe1,text = "",font = ("Helvetica",8,"italic"))
        self.name_entry.grid(row = 1, column = 0,sticky = "w",padx = 37)

        self.surname = Label(self.newframe1, text="Surname",font = ("Helvetica",9,"bold"))
        self.surname.grid(sticky=W, row=2, column=0, pady=7)

        self.surname_entry = Label(self.newframe1,text = "",font = ("Helvetica",8,"italic"))
        self.surname_entry.grid(row=2, column=0, sticky="w", padx=60)

        self.ID = Label(self.newframe1, text="ID",font = ("Helvetica",9,"bold"))
        self.ID.grid(sticky=W, row=3, column=0)

        self.ID_entry = Label(self.newframe1, text="",font = ("Helvetica",9,"italic"))
        self.ID_entry.grid(row=3, column=0, sticky="w", padx=60)

        self.Dept = Label(self.newframe1, text="Dept",font = ("Helvetica",9,"bold"))
        self.Dept.grid(sticky=W, row=4, column=0, pady=7)

        self.Dept_entry = Label(self.newframe1, text="",font = ("Helvetica",9,"italic"))
        self.Dept_entry.grid(row=4, column=0, sticky="w", padx=60)

        self.GPA = Label(self.newframe1, text="GPA",font = ("Helvetica",9,"bold"))
        self.GPA.grid(sticky=W, row=5, column=0)

        self.GPA_entry = Label(self.newframe1, text="",font = ("Helvetica",9,"italic"))
        self.GPA_entry.grid(row=5, column=0, sticky="w", padx=60)



        self.MP1 = Label(self.newframe1, text="MP1 Grade:",font = ("Helvetica",9,"bold"))
        self.MP1.grid(row=1, column=0,sticky = "E",padx = 100)

        self.MP1_grade = Label(self.newframe1, text="",  font=("Helvetica", 9, "italic"))
        self.MP1_grade.grid(row=1 ,sticky = "NE",padx = 60,pady = 10)

        self.MP2 = Label(self.newframe1, text="MP2 Grade:",font = ("Helvetica",9,"bold"))
        self.MP2.grid(row=2, column=0,sticky = "E",padx = 100)

        self.MP2_grade = Label(self.newframe1, text="", font=("Helvetica", 9, "italic"))
        self.MP2_grade.grid(row=2, sticky="NE", padx=60, pady=10)

        self.MP3 = Label(self.newframe1, text="MP3 Grade:",font = ("Helvetica",9,"bold"))
        self.MP3.grid(row=3, column=0,sticky = "E",padx = 100)

        self.MP3_grade= Label(self.newframe1, text="", font=("Helvetica", 9, "italic"))
        self.MP3_grade.grid(row=3, sticky="NE", padx=60, pady=10)

        self.MT = Label(self.newframe1, text="MT Grade:",font = ("Helvetica",9,"bold"))
        self.MT.grid(row=4, column=0,sticky = "E",padx = 100)

        self.MT_grade= Label(self.newframe1, text="", font=("Helvetica", 9, "italic"))
        self.MT_grade.grid(row=4, sticky="NE", padx=60, pady=10)

        self.Final = Label(self.newframe1, text="Final Grade:",font = ("Helvetica",9,"bold"))
        self.Final.grid(row=5, column=0,sticky = "E",padx = 100)

        self.Final_grade= Label(self.newframe1, text="", font=("Helvetica", 9, "italic"))
        self.Final_grade.grid(row=5, sticky="NE", padx=60, pady=10)

        self.newframe2 = Frame(self, borderwidth=2, relief=GROOVE)
        self.newframe2.grid(sticky=W, pady=3)                       #Frame 2 will hold frame 4

        self.newframe4 = Frame(self.newframe2, relief=GROOVE) #frame that will hold the bottom (labels/entries of grades,and other labels and button)
        self.newframe4.grid(sticky=W, row=0, column=0)
        self.projects = Label(self.newframe4, text="Projects:")
        self.projects.grid(sticky=W, row=0, column=0)

        Label(self.newframe4, text="MP1").grid(row=0, column=1)
        Label(self.newframe4, text="MP2").grid(row=0, column=2)
        Label(self.newframe4, text="MP3").grid(row=0, column=3)
        Label(self.newframe4, text="MT").grid(row=0, column=4)
        Label(self.newframe4, text="Final").grid(row=0, column=5)
        self.projects = Label(self.newframe4, text="Grades:", borderwidth=2)
        self.mp1_entry = Entry(self.newframe4, textvariable=self.first_grade_entry, width=9).grid(row=1, column=1, padx=3)
        self.mp2_entry=Entry(self.newframe4, textvariable=self.second_grade_entry, width=9).grid(row=1, column=2, padx=3)
        self.mp3_entry=Entry(self.newframe4, textvariable=self.third_grade_entry, width=9).grid(row=1, column=3, padx=3)
        self.mt_entry=Entry(self.newframe4, textvariable=self.forth_grade_entry, width=9).grid(row=1, column=4, padx=3)
        self.final_entry=Entry(self.newframe4, textvariable=self.fifth_grade_entry, width=9).grid(row=1, column=5, padx=3)
        self.projects.grid(sticky=W, row=1, column=0)
        self.projects = Label(self.newframe4, text="Export As:")
        self.projects.grid(sticky=W, row=2, column=0)

        self.Save_Grades_Button = Button(self.newframe4, text="Save Grades",font = ("Helvetica",9,"bold"),command = self.save_grades)
        self.Save_Grades_Button.grid(row=1, column=6, padx=29)
        self.space = Label(self.newframe2, width=23, text=" ").grid(row=1, column=7)

        self.newframe3 = Frame(self.newframe2, relief=GROOVE) #Frame to hold check boxes and export data stuff)
        self.newframe3.grid(row=5, column=0, sticky="W", padx=80)
        self.Checkbox1 = Checkbutton(self.newframe3, text='csv', variable=self.Checkbox, onvalue='csv').grid(row=0,
                                                                                                             column=0,sticky="w",padx=0)
        self.Checkbox2 = Checkbutton(self.newframe3, text='txt', variable=self.Checkbox, onvalue='txt').grid(row=1,column=0,padx=0,sticky="w")
        self.Checkbox3 = Checkbutton(self.newframe3, text='xls', variable=self.Checkbox, onvalue='xls').grid(row=2,column=0,padx=0,sticky="w")

        self.File_Name_Label = Label(self.newframe3, text="File Name:").grid(row=0, column=2, sticky="W")
        self.File_Name_Entry = Entry(self.newframe3, width=20,textvariable = self.var_of_filename).grid(row=0,column=2,sticky="e")
        self.Export_Data_Button = Button(self.newframe3, text="Export Data", width=25,command = self.export_file).grid(row=1, column=2)

        self.message = Label(self, text="Program messages...", fg = "red")
        self.message.grid(row=7, sticky="w")
        self.pack()


    def file_selection(self): #Function of the select file button
        self.message.config(text="INFO: File Selection.")
        try:#calls second class reading excel function and and inserts to tree widget
            self.StudentList.reading_excel()
            self.message.config(text="INFO: File Loaded.")
            self.control_variable = 1   #variable changed to when to indicate that file has been imported successfully
            for i in self.StudentList.Students: #inserting data in the tree widget
                try:#if student has two names and one surname
                    self.Treeview.insert("", 'end', values=[int(self.StudentList.Students[i].ID),
                                                            self.StudentList.Students[i].Name[0] + " " +
                                                            self.StudentList.Students[i].Name[1],
                                                            self.StudentList.Students[i].Name[2]])
                except:# if student has one name and one surname
                    self.Treeview.insert("", 'end', values=[int(self.StudentList.Students[i].ID),
                                                            self.StudentList.Students[i].Name[0],
                                                            self.StudentList.Students[i].Name[1]])
        except xlrd.XLRDError:#if file not accepted(failed to load)
            self.message.config(text="INFO: Loading Failed! Please Try Again Later.", fg="Red")
        except IOError:#if user cancels file selection process
            self.message.config(text = "INFO: File Selection Canceled.")
        except IndexError:
            self.message.config(text = "INFO: Excel File Has Different Format.")
        except:
            self.message.config(text="INFO: Loading Failed! Please Try Again Later.")







    def show_data(self):#function of show data button

        if (self.Treeview.selection() == ()) and (self.control_variable == 1):#condition to check if student is selected and file is imported
            self.message.config(text="INFO: Please Select A Student First.", fg="Red")
        elif self.control_variable == 0:#condition to check if file is imported or not
            self.message.config(text="INFO: Please Load The Files First.", fg="Red")

        for i in self.Treeview.selection():#adding the selected student info to the student details frame
            self.dict_of_row = self.Treeview.item(i)
            self.list_of_values = self.dict_of_row['values']
            self.name_entry.config(text=self.list_of_values[1])
            self.surname_entry.config(text=self.list_of_values[2])
            self.ids = str(self.list_of_values[0]).split(".")[0]
            self.ID_entry.config(text=str(self.ids))
            self.student = self.StudentList.Students[self.ids]
            self.Dept_entry.config(text=self.student.Dept)
            self.GPA_entry.config(text=str(self.student.GPA))
            self.MP1_grade.config(text=str(self.student.MP1))
            self.MP2_grade.config(text=str(self.student.MP2))
            self.MP3_grade.config(text=str(self.student.MP3))
            self.MT_grade.config(text=str(self.student.MT))
            self.Final_grade.config(text=str(self.student.FINAL))
            self.first_grade_entry.set(self.student.MP1)
            self.second_grade_entry.set(self.student.MP2)
            self.third_grade_entry.set(self.student.MP3)
            self.forth_grade_entry.set(self.student.MT)
            self.fifth_grade_entry.set(self.student.FINAL)
    def save_grades(self):#function of save grade button

        try:#tries saving and fails if student isnt selected or file isnt imported
            non_accepted_value = 0 #variable that increases if mp1,2,3,mt,final entry has any non accepted values
            for i in self.first_grade_entry.get():#iterates through every letter in the input and checks if it is allowed
                if i not in ["0","1","2","3","4","5","6","7","8","9",'""']:
                    self.message.config(text="INFO: Warning, The Type Of The Grade Is Incorrect.")
                    non_accepted_value+=1

            if non_accepted_value == 0:#if it doesnt increase it updates the grade                   ---------ALL FOLLOWING LINES OF CODE ARE THE SAME FOR DIFFERENT ENTRIES------------------
                self.student.MP1 = self.first_grade_entry.get()
                self.message.config(text = "INFO: Grades Successfully Saved.")

                non_accepted_value = 0
            for i in self.second_grade_entry.get():
                    if i not in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", '""']:
                        self.message.config(text="INFO: Warning, The Type Of The Grade Is Incorrect.")
                        non_accepted_value+=1
            if non_accepted_value == 0:
                self.student.MP2 = self.second_grade_entry.get()

                non_accepted_value = 0
            for i in self.third_grade_entry.get():
                    if i not in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", '""']:
                        self.message.config(text="INFO: Warning, The Type Of The Grade Is Incorrect.")
                        non_accepted_value+=1
            if non_accepted_value == 0:
                self.student.MP3 = self.third_grade_entry.get()

                non_accepted_value = 0
            for i in self.forth_grade_entry.get():
                    if i not in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", '""']:
                        self.message.config(text="INFO: Warning, The Type Of The Grade Is Incorrect.")
                        non_accepted_value+=1
            if non_accepted_value == 0:
                self.student.MT = self.forth_grade_entry.get()

                non_accepted_value = 0
            for i in self.fifth_grade_entry.get():
                    if i not in ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", '""']:
                        self.message.config(text="INFO: Warning, The Type Of The Grade Is Incorrect.")
                        non_accepted_value+=1
            if non_accepted_value == 0:
                self.student.FINAL = self.fifth_grade_entry.get()
            self.show_data()#Displays data of student



        except:#if the try method fails it checks if its because file isnt imported or student isnt selected

            if self.control_variable == 0:
               self.message.config(text = "INFO: Please Load The Files Fisrt.",fg = "Red")
            elif self.Treeview.selection() == ():
                self.message.config(text = "INFO: Please Select A Student First. ",fg = "Red")
            else:
                self.message.config(text = "INFO: Select Show Data First.")



    def export_file(self): #fucntion for export data button

        if self.Checkbox.get() == "xls": # if user chooses excel file type
            book = Workbook() #creates excel
            sheet1 =  book.add_sheet('Sheet 1',cell_overwrite_ok=True)  #adds sheet

            sheet1.write(0, 0, 'ID')            #creating the first defined row with a defined width
            sheet1.write(0, 1, 'NAME')
            sheet1.write(0, 2, 'SECTION')
            sheet1.write(0, 3, 'DEPT')
            sheet1.write(0, 4, 'GPA')
            sheet1.write(0, 5, 'MP1')
            sheet1.write(0, 6, 'MP2')
            sheet1.write(0, 7, 'MP3')
            sheet1.write(0, 8, 'MT')
            sheet1.write(0, 9, 'FINAL')
            sheet1.col(0).width = 3500
            sheet1.col(1).width = 3500
            sheet1.col(2).width = 3500
            sheet1.col(3).width = 3500
            sheet1.col(4).width = 3500
            sheet1.col(5).width = 3500
            sheet1.col(6).width = 3500
            sheet1.col(7).width = 3500
            sheet1.col(8).width = 3500
            sheet1.col(9).width = 3500
            row = 0   #variable that increases to change row
            for child in self.Treeview.get_children():
                row+=1  #increasing row

                sheet1.write(row,0,self.Treeview.item(child)["values"][0])#-------ADDING THE DIFFERENT STUDENT DETAILS TO DIFFERENT COLUMNS -------------
                sheet1.write(row,1,self.Treeview.item(child)["values"][1]+" "+self.Treeview.item(child)["values"][2])
                sheet1.write(row,2,self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].Section)
                sheet1.write(row,3, self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].Dept)
                sheet1.write(row, 4, self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].GPA)
                sheet1.write(row, 5, self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MP1,easyxf('align : horiz right'))
                sheet1.write(row, 6, self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MP2,easyxf('align : horiz right'))
                sheet1.write(row, 7, self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MP3,easyxf('align : horiz right'))
                sheet1.write(row, 8, self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MT,easyxf('align : horiz right'))
                sheet1.write(row, 9, self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].FINAL,easyxf('align : horiz right'))
            if self.var_of_filename.get().count(" ")== len(self.var_of_filename.get()):#CONDITION TO WARN USER IF FILE NAME IS NOT PROVIDED
                self.message.config(text="INFO: Please Provide The Name Of The File.")

            else: #IF ABOVE CONDITION IS NOT MET IT WILL NAME\SAVE THE FILE
                try: book.save(self.var_of_filename.get() + '.xls'),self.message.config(text = "INFO: File Saved As "+"'"+str(self.var_of_filename.get())+".xls'.")
                except IOError: self.message.config(text = "INFO: File Already In Use, Please Close The File " + "'"+str(self.var_of_filename.get()) + ".xls'.")
        elif self.Checkbox.get() == "csv":#IF FILE TYPE IS CSV IT RAISES A WARNING
            self.message.config(text = "INFO: Type Not Supported.")
        elif self.Checkbox.get() == "e":#IF FILE TYPE VARIABLE IS STILL THE DEFAULT THEN IN WARNS USER TO SELECT FILE TYPE
            self.message.config(text = "INFO: Type Not Chosen.")
        elif self.Checkbox.get() == "txt":#If File type is txt
            if self.var_of_filename.get().count(" ")== len(self.var_of_filename.get()):#if name is not provided user is warned
                self.message.config(text="INFO: Please Provide The Name Of The File.")
            else:
                file1 = open(self.var_of_filename.get()+".txt","w")#opens txt file in write mode
                for child in self.Treeview.get_children():#iterates over each student in tree widget
                    file1.write(str(self.Treeview.item(child)["values"][0])+" , "+self.Treeview.item(child)["values"][1].encode("utf8")+" , "+#adds each detail of the student to one line then creates a new line
                                self.Treeview.item(child)["values"][2].encode("utf8")+" , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].Section)+
                                " , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].Dept)+" , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].GPA)+
                                " , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MP1)+" , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MP2)+
                                " , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MP3)+" , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].MT)+
                                " , "+str(self.StudentList.Students[str(self.Treeview.item(child)["values"][0])].FINAL)+".\n")
                file1.close()
                self.message.config(text = "INFO: File Saved As: "+"'"+str(self.var_of_filename.get())+".txt'.")






def main():# function that runs the app by creating object of tk and third class which takes the object of tk as argument
    root = Tk()
    root.title('Grades Management Tool v1.0')
    root.geometry('690x525')
    app = GUI(root)
    root.mainloop()


main()
