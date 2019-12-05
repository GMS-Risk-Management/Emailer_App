import tkinter as tk
from tkinter import ttk, StringVar, filedialog, font, Frame
from dateutil.parser import parse
import numpy as np
from pandas import read_excel, Series, DataFrame
import mammoth
import win32com.client as win32
import pandas as pd

class Emailer:


    def __init__(self, parent):

        font_type = font.Font(family='Helvetica', size=12, weight='bold')
        font_type_2 = font.Font(family='Helvetica', size=10)

        def set_remaining_count(item):
            current_count = len(self.sent_list)
            if item in self.sent_list:
                self.sent_list.remove(item)
                current_count = len(self.sent_list)
                self.email_button.configure(text='Create Emails: Number New of Unique GPNS (' + str(current_count) + ')')
            else:
                pass
            return None

        def send_notification():
            gpn_list = self.gpn_choice['values']
            gpn_list = list(set(list(gpn_list)))
            gpn_list = gpn_list[int(self.start.get())-1:int(self.end.get())]
            #test code
            subject = self.subject_line.get("1.0","end")
            #end test code
            for item in gpn_list:
                recipient = item
                message = create_table(item)
                message = special_formatting(message,item)
                # message = self.html_output.get("1.0",'end')
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.To = recipient
                mail.Subject = subject
                mail.HTMLBody = message
                mail.Importance = 2
                mail.Display()
                set_remaining_count(item)
                mail.Send()
            return None

        def get_email_from_word():
            root = tk.Tk()
            file = tk.filedialog.askopenfilenames(parent=root,title='Word Document?')[0]
            root.destroy()
            with open(file, "rb") as docx_file:
                result = mammoth.convert_to_html(docx_file)
                html = result.value # The generated HTML

            self.html_output.delete('1.0', tk.END)
            self.html_output.insert('1.0', str(html))
            self.html_output.tag_configure("center", justify='center')
            self.html_output.tag_add("center", "1.0", "end")
            self.html_output.configure(font=("Helvetica Neue", 10))
            #Insert code to grab word document and use mammoth to convert to html
            return None


        def get_gpn_field_list():
            root = tk.Tk()
            file = tk.filedialog.askopenfilenames(parent=root,title='Excel File?')[0]
            root.destroy()
            data = read_excel(file)
            gpns = data['GPN']
            count_gpn = len(set(gpns.tolist()))
            fields = data.columns
            self.gpn_choice['values'] = gpns.dropna().tolist()
            self.sent_list = list(set(gpns.dropna().tolist()))
            self.field_choice['values'] = fields.dropna().tolist()
            self.data_frame = data
            self.email_button.configure(text='Create Emails: Number of New Unique GPNS (' + str(count_gpn) + ')')
            return None


        def select_fields(event):
            selection = self.field_choice.get()
            self.field_output.insert("end",selection + ',')
            return None

        def clear_fields():
            self.field_output.delete('1.0', tk.END)
            return None

        def clear_subject():
            self.subject_line.delete('1.0', tk.END)
            return None

        def clear_formatting():
            self.formatting_box.delete('1.0', tk.END)
            return None

        def create_table(gpn):
            text = self.field_output.get('1.0','end')
            if text.find(',')==-1:
                message = self.html_output.get('1.0','end')
            else:
                pd.set_option('display.max_colwidth', -1)
                arr = self.field_output.get("1.0","end")
                arr = arr[0:arr.rfind(',')]
                arr = arr.split(',')
                data = self.data_frame.loc[self.data_frame['GPN']==gpn]
                data = data[arr]
                data = data.to_html(index=False,justify='center',col_space=50).replace("\\n","<br>")
                message = self.html_output.get("1.0","end")
                message = message.replace('{table}',data)
            return message

        def special_formatting(message,gpn):
            if self.formatting_box.compare("end-1c", "==", "1.0"):
                return message
            else:
                arr = self.formatting_box.get("1.0","end")
                print(arr)
                arr = arr.replace('\n','')
                arr = arr.split(',')
                print(arr)
                for item in arr:
                    message = message.replace('{'+item+'}',str(self.data_frame.loc[self.data_frame['GPN']==gpn][item].to_list()[0]))
            return message

        self.data_frame = DataFrame()
        self.sent_list = []

        self.parent = parent
        today = str(parse(str(np.datetime64('today'))))[:10]
        self.lbl = tk.Label(self.parent, text="Automatic Emailer Application - " + today)
        self.lbl.grid(row=0,column=0,columnspan=6,padx=5,pady=5,sticky='nesw')
        self.lbl.configure(font=font_type)
        self.lbl.configure(background='#f8d21c')

        self.lbl = tk.Label(self.parent, text="Start:")
        self.lbl.grid(row=4,column=0,padx=5,columnspan=1,sticky='nesw')
        self.lbl.configure(font=font_type_2)
        self.lbl.configure(background='#f8d21c')

        self.lbl = tk.Label(self.parent, text="End:")
        self.lbl.grid(row=4,column=3,padx=5,columnspan=1,sticky='nesw')
        self.lbl.configure(font=font_type_2)
        self.lbl.configure(background='#f8d21c')

        self.start = tk.Entry(self.parent,justify='center')
        self.start.grid(row=4,column=2,padx=5,columnspan=1,sticky='nesw')

        self.end = tk.Entry(self.parent,justify='center')
        self.end.grid(row=4,column=5,padx=5,columnspan=1,sticky='nesw')

        self.word_button = tk.Button(self.parent, text="Select Word Document", command=get_email_from_word)
        self.word_button.grid(row=1,column=0,padx=5,pady=5,columnspan=6,sticky='nesw')
        self.word_button.configure(foreground='#f8d21c')
        self.word_button.configure(background='#333')
        self.word_button.configure(font=font_type_2)

        self.gpn_button = tk.Button(self.parent, text="Select Excel File", command=get_gpn_field_list)
        self.gpn_button.grid(row=2,column=0,padx=5,pady=5,columnspan=6,sticky='nesw')
        self.gpn_button.configure(foreground='#f8d21c')
        self.gpn_button.configure(background='#333')
        self.gpn_button.configure(font=font_type_2)

        self.email_button = tk.Button(self.parent, text="Create Emails", command=send_notification)
        self.email_button.grid(row=3,column=0,columnspan=6,padx=5,pady=5,sticky='nesw')

        self.email_button.configure(foreground='#f8d21c')
        self.email_button.configure(background='#333')
        self.email_button.configure(font=font_type_2)

        self.gpn_choice = ttk.Combobox(self.parent,
                        state='readonly',
                        width = 30
                        )

        self.gpn_choice['values'] = ['GPN List']
        self.gpn_choice.current(0)
        self.gpn_choice.grid(row=5,column=0,columnspan=6,pady=5)
        self.gpn_choice.configure(font=font_type_2)

        self.field_choice = ttk.Combobox(self.parent,
                        state='readonly',
                        width = 30
                        )

        self.field_choice['values'] = ['Field List']
        self.field_choice.current(0)
        self.field_choice.grid(row=7,column=0,columnspan=6,pady=5)
        self.field_choice.bind('<<ComboboxSelected>>', select_fields)
        self.field_choice.configure(font=font_type_2)


        self.html_output = tk.Text(self.parent,
                                    height=15,
                                    width=80,
                                    wrap='word'
                                        )

        self.html_output.grid(row=0,column=7,padx=5,pady=5,rowspan=6,columnspan=6,sticky='nesw')

        self.html_output.insert('1.0','\n\n\n\n' + 'No Message Text Currently Selected')
        self.html_output.tag_configure("center", justify='center')
        self.html_output.tag_add("center", "1.0", "end")
        self.html_output.configure(font=("Helvetica Neue", 10))

        self.lbl = tk.Label(self.parent, text="Fields for Table")
        self.lbl.grid(row=7,column=7,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')
        self.lbl.configure(font=font_type_2)
        self.lbl.configure(background='#f8d21c')

        self.lbl = tk.Label(self.parent, text="Subject Line")
        self.lbl.grid(row=7,column=9,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')
        self.lbl.configure(font=font_type_2)
        self.lbl.configure(background='#f8d21c')

        self.lbl = tk.Label(self.parent, text="Special Formattng")
        self.lbl.grid(row=7,column=11,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')
        self.lbl.configure(font=font_type_2)
        self.lbl.configure(background='#f8d21c')

        self.field_output = tk.Text(self.parent,
                                    height=5,
                                    width=20,
                                    wrap='word'
                                        )

        self.field_output.grid(row=8,column=7,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')

        self.field_output.tag_configure("center", justify='center')
        self.field_output.tag_add("center", "1.0", "end")
        self.field_output.configure(font=("Helvetica Neue", 10))

        self.subject_line = tk.Text(self.parent,
                                    height=5,
                                    width=20,
                                    wrap='word'
                                        )

        self.subject_line.grid(row=8,column=9,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')

        self.subject_line.tag_configure("center", justify='center')
        self.subject_line.tag_add("center", "1.0", "end")
        self.subject_line.configure(font=("Helvetica Neue", 10))

        self.formatting_box = tk.Text(self.parent,
                                    height=5,
                                    width=20,
                                    wrap='word'
                                        )

        self.formatting_box.grid(row=8,column=11,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')

        self.formatting_box.tag_configure("center", justify='center')
        self.formatting_box.tag_add("center", "1.0", "end")
        self.formatting_box.configure(font=("Helvetica Neue", 10))

        self.clear_button = tk.Button(self.parent, text="Clear Fields", command=clear_fields)
        self.clear_button.grid(row=9,column=7,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')

        self.clear_button.configure(foreground='#f8d21c')
        self.clear_button.configure(background='#333')
        self.clear_button.configure(font=font_type_2)

        self.clear_subject = tk.Button(self.parent, text="Clear Subject", command=clear_subject)
        self.clear_subject.grid(row=9,column=9,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')

        self.clear_subject.configure(foreground='#f8d21c')
        self.clear_subject.configure(background='#333')
        self.clear_subject.configure(font=font_type_2)

        self.clear_formatting = tk.Button(self.parent, text="Clear Formatting", command=clear_formatting)
        self.clear_formatting.grid(row=9,column=11,padx=5,pady=5,rowspan=1,columnspan=2,sticky='nesw')

        self.clear_formatting.configure(foreground='#f8d21c')
        self.clear_formatting.configure(background='#333')
        self.clear_formatting.configure(font=font_type_2)


if __name__ == '__main__':
    today = str(parse(str(np.datetime64('today'))))[:10]
    root = tk.Tk()
    root.title('GMS Automatic Emailer - ' + 'Triantaphilos Manning Testing')
    root.geometry('940x430')
    root.configure(background='#333')
    # photo = PhotoImage(file = r'Z:\Policy Finder\InkedDiscover_LI.gif')
    # root.iconphoto(False, photo)
    app = Emailer(root)
    root.mainloop()
