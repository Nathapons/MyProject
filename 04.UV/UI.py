from settings import ReadProgram
from tkinter import messagebox, ttk
from tkinter import *
from tkinter.filedialog import asksaveasfilename
import pandas as pd
import webbrowser
import os


class UvProgram(ReadProgram):
    def __init__(self):
        super().__init__()
        
    def userform(self):
        root = Tk()
        WIDTH = 700
        HEIGHT = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        center_width = int((screen_width/2) - (WIDTH/2))
        center_height = int((screen_height/2) - (HEIGHT/1.8))
        # Root Properties
        root.geometry(f'{WIDTH}x{HEIGHT}+{center_width}+{center_height}')
        root.title('UV Program')
        root.config(bg='white')
        root.resizable(0, 0)
        # Widget
        big_frame = Frame(root, bg='white')
        big_frame.pack()
        topic = Label(big_frame, text='UV Program', font=('Arial', 24), fg='white', bg='blue')
        topic.grid(row=0, column=0, pady=10, ipadx=10)
        cb_frame = Frame(big_frame, bg='white')
        cb_frame.grid(row=1, column=0)
        self.__string_unit = StringVar(value='mW/cm2')
        self.machine = StringVar()
        machine_label = Label(cb_frame, text='MACHINE:', font=('Arial', 14), bg='white')
        machine_label.grid(row=0, column=0, padx=5, pady=5)
        machine_entry = Entry(cb_frame, font=('Arial', 14), bg='white', width=16, justify='center', textvariable=self.machine)
        machine_entry.grid(row=0, column=1, padx=5, pady=5)
        target_label = Label(cb_frame, text='TARGET:', font=('Arial', 14), bg='white')
        target_label.grid(row=0, column=2, padx=5, pady=5)
        self.target = StringVar(value='30')
        target_val = 30
        target_list = []
        while target_val <= 160:
            if target_val != 140:
                target_list.append(str(target_val))
            target_val += 10

        self.target_cb = ttk.Combobox(
              cb_frame
            , background='#ffff66'
            , font=('Arial', 14)
            , values=target_list
            , textvariable=self.target
            , justify='center'
            , width=14
        )
        self.target_cb.grid(row=0, column=3, padx=5, pady=5)
        settings_label = Label(cb_frame, text='SETTINGS:', font=('Arial', 14), bg='white')
        settings_label.grid(row=1, column=0, padx=5, pady=5)
        self.settings = StringVar()
        settings_entry = Entry(cb_frame, font=('Arial', 14), bg='white', width=16, justify='center', textvariable=self.settings)
        settings_entry.grid(row=1, column=1, padx=5, pady=5)
        self.hrs = StringVar(value='0')
        hrs_label = Label(cb_frame, text='HOURS:', font=('Arial', 14), bg='white')
        hrs_label.grid(row=1, column=2, padx=5, pady=5)
        self.hrs_cb = ttk.Combobox(
              cb_frame
            , background='#ffff66'
            , font=('Arial', 14)
            , values=['0', '150', '300', '450', '600', '750', '900', '1050', '1200', '1350', '1500', '1650']
            , textvariable=self.target
            , justify='center'
            , width=14
        )
        self.hrs_cb.grid(row=1, column=3, padx=5, pady=5)
        self.side = StringVar(value='F')
        side_label = Label(cb_frame, text='SIDE:', font=('Arial', 14), bg='white')
        side_label.grid(row=2, column=0, padx=5, pady=5)
        self.side_cb = ttk.Combobox(
              cb_frame
            , background='#ffff66'
            , font=('Arial', 14)
            , values=['F', 'B']
            , textvariable=self.side
            , justify='center'
            , width=14
        )
        self.side_cb.grid(row=2, column=1, padx=5, pady=5)
        self.position = StringVar(value='1')
        position = Label(cb_frame, text='POSITION:', font=('Arial', 14), bg='white')
        position.grid(row=2, column=2, padx=5, pady=5)
        self.side_cb = ttk.Combobox(
              cb_frame
            , background='#ffff66'
            , font=('Arial', 14)
            , values=[i+1 for i in range(10)]
            , textvariable=self.position
            , justify='center'
            , width=14
        )
        self.side_cb.grid(row=2, column=3, padx=5, pady=5)
        unit_label = Label(cb_frame, text='UNIT:', font=('Arial', 14), bg='white')
        unit_label.grid(row=3, column=0, padx=5, pady=5)
        self.unit_cb = ttk.Combobox(
              cb_frame
            , background='#ffff66'
            , font=('Arial', 14)
            , values=['mW/cm2', 'mJ/cm2']
            , textvariable=self.__string_unit
            , justify='center'
            , width=14
        )
        self.unit_cb.grid(row=3, column=1, padx=10)
        s = ttk.Style()
        s.configure('my.TButton', font=('Arial', 14))
        button_frame = Frame(big_frame, bg='white')
        button_frame.grid(row=2, column=0, pady=10)
        self.run_program = ttk.Button(
               button_frame
            , text='Record Value'
            , style='my.TButton'
        )
        self.run_program.grid(row=0, column=1, pady=5, ipady=2)
        self.run_program.bind('<Button-1>', self.check_condition)
        self.csv_program = ttk.Button(
               button_frame
            , text='Export Excel'
            , style='my.TButton'
        )
        self.csv_program.grid_forget()
        self.csv_program.bind('<Button-1>', self._export_csv)

        self.headers = ['NO.', 'MACHINE', 'UV-TARGET', 'UV-SETTING', 'HOURS', 'SIDE', 'POSTION','ACTUAL', 'UNIT']
        treeview_frame = Frame(big_frame, bg='white')
        treeview_frame.grid(row=3, column=0, pady=10)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
        self.__status_tree = ttk.Treeview(treeview_frame, column=self.headers, show='headings', height=13)
        for header in self.headers:
            if header in ['SIDE', 'POSITION', 'HOURS', 'ACTUAL', 'UNIT']:
                col_width = 60
            elif header == 'NO.':
                col_width = 30
            elif header == 'MACHINE':
                col_width = 100
            else:
                col_width = 90
            self.__status_tree.heading(header, text=header)
            
            self.__status_tree.column(header, anchor='center', width=col_width, minwidth=0)
        self.__status_tree.grid(row=0, column=0)
        vertical_scrollbar = ttk.Scrollbar(treeview_frame, orient="vertical", command=self.__status_tree.yview)
        vertical_scrollbar.grid(row=0, column=1, ipady=100)
        root.mainloop()

    def check_condition(self, event):
        if self.machine.get() == "" and self.settings.get() == "":
            messagebox.showwarning(
                title="Alarm to users"
                , message="กรุณากรอกข้อมูลที่ช่อง Machine และ Settings ก่อนบันทึกค้า"
            )
            return
        # Call Super class from ReadProgram
        port = super().get_ports()

        # Check Port is connect
        if port == 'No Connect':
            messagebox.showwarning(
                    title='Warning'
                , message='กรุณาเช็คสาย UV ว่าเชื่อมต่อกับ PC หรือไม่?'
            )
            self.run_program['text'] = 'Record Value'
        else:
            # Run Function to get treeview
            self._update_treeview()
            # Count data in treeview
            total_add = len(self.__status_tree.get_children())
            if total_add != 0:
                self.csv_program.grid(row=0, column=0, padx=5, ipady=2)

    def _update_treeview(self):
        num = len(self.__status_tree.get_children()) + 1
        machine = self.machine.get()
        uv_target = self.target_cb.get()
        uv_settings = self.settings.get()
        hours = self.hrs_cb.get()
        side = self.side.get()
        position = self.position.get()
        actual = super().read_value()
        str_unit = self.__string_unit.get()
        data = [
              num
            , machine
            , uv_target
            , uv_settings
            , hours
            , side
            , position
            , actual
            , str_unit
        ]

        self.__status_tree.insert('', 'end', values=data)

    def _export_csv(self, event):
        total_add = len(self.__status_tree.get_children())
        if total_add > 0:
            data = [self.__status_tree.item(item)['values'] for item in self.__status_tree.get_children()]
            df = pd.DataFrame(data, columns= self.headers)

            messagebox.showinfo(title="Information", message="กรุณาตั้งชื่อไฟล์ Result")
            try:
                # csv_file = asksaveasfilename(filetypes=[("CSV files", '*.csv')], defaultextension='.csv')
                excel_file = asksaveasfilename(filetypes=[("Excel files", '*.xlsx')], defaultextension='.xlsx')
                # df.to_csv(csv_file, index = False, header=True)
                writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='UV', index=False, header=True)
                writer.save()

                self._delete_treeview()
                answer = messagebox.askyesno(title='Ask to User', message=f'คุณต้องการจะเปิดไฟล์ {os.path.basename(excel_file)}')
                self.csv_program.grid_forget()
                if answer:
                    webbrowser.open(excel_file)

            except:
                messagebox.showwarning(title="Information", message="คุณตั้งชื่อไฟล์ผิดกรุณาลองใหม่อีกครั้ง")
            self.csv_program.grid_forget()

    def _delete_treeview(self):
        for i in self.__status_tree.get_children():
            self.__status_tree.delete(i)