from settings import ReadProgram
from tkinter import messagebox, ttk
from tkinter import *
from tkinter.filedialog import asksaveasfilename
import pandas as pd
import webbrowser
import os


class UvProgram():
    def userform(self):
        root = Tk()
        WIDTH = 450
        HEIGHT = 400
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
        self.string_unit = StringVar(value='mW/cm2')
        unit_label = Label(cb_frame, text='UNIT:', font=('Arial', 14), bg='white')
        unit_label.grid(row=0, column=0, padx=10)
        self.unit_cb = ttk.Combobox(
              cb_frame
            , background='#ffff66'
            , font=('Arial', 14)
            , values=['mW/cm2', 'mJ/cm2']
            , textvariable=self.string_unit
            , justify='center'
            , width=15
        )
        self.unit_cb.grid(row=0, column=1, padx=10)
        s = ttk.Style()
        s.configure('my.TButton', font=('Arial', 14))
        button_frame = Frame(big_frame, bg='white')
        button_frame.grid(row=2, column=0, pady=10)
        self.run_program = ttk.Button(
               button_frame
            , text='Run Program'
            , style='my.TButton'
        )
        self.run_program.grid(row=0, column=1, pady=5, ipady=2)
        self.run_program.bind('<Button-1>', self.check_port_open)
        self.csv_program = ttk.Button(
               button_frame
            , text='Export CSV'
            , style='my.TButton'
        )
        self.csv_program.grid_forget()
        self.csv_program.bind('<Button-1>', self.export_csv)

        headers = ['NO.', 'RESULT', 'UNIT']
        treeview_frame = Frame(big_frame, bg='white')
        treeview_frame.grid(row=3, column=0, pady=10)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
        self.status_tree = ttk.Treeview(treeview_frame, column=headers, show='headings', height=9)
        for header in headers:
            self.status_tree.heading(header, text=header)
            col_width = 120
            self.status_tree.column(header, anchor='center', width=col_width, minwidth=0)
        self.status_tree.grid(row=0, column=0)
        vertical_scrollbar = ttk.Scrollbar(treeview_frame, orient="vertical", command=self.status_tree.yview)
        vertical_scrollbar.grid(row=0, column=1, ipady=80)

        root.mainloop()

    def check_port_open(self, event):
        self.obj = ReadProgram()
        port = self.obj.get_ports()

        # Check Port is connect
        if port == 'No Connect':
            messagebox.showwarning(
                    title='Warning'
                , message='กรุณาเช็คสาย UV ว่าเชื่อมต่อกับ PC หรือไม่?'
            )
            self.run_program['text'] = 'Run Program'
            quit()
        else:
            # Run Function to get treeview
            self.update_treeview()
            # Count data in treeview
            total_add = len(self.status_tree.get_children())
            if total_add != 0:
                self.csv_program.grid(row=0, column=0, padx=5, ipady=2)

    def update_treeview(self):
        num = len(self.status_tree.get_children()) + 1
        value = self.obj.read_value()
        data = [num, value]
        self.status_tree.insert('', 'end', values=data)

    def export_csv(self, event):
        total_add = len(self.status_tree.get_children())
        if total_add > 0:
            data = [self.status_tree.item(item)['values'] for item in self.status_tree.get_children()]
            df = pd.DataFrame(data, columns= ['NO.', 'RESULT', 'UNIT'])

            messagebox.showinfo(title="Information", message="กรุณาตั้งชื่อไฟล์ Result")
            try:
                csv_file = asksaveasfilename(filetypes=[("CSV files", '*.csv')], defaultextension='.csv')
                df.to_csv(csv_file, index = False, header=True)

                self.delete_treeview()
                answer = messagebox.askyesno(title='Ask to User', message=f'คุณต้องการจะเปิดไฟล์ {os.path.basename(csv_file)}')
                if answer:
                    webbrowser.open(csv_file) 
            except:
                messagebox.showwarning(title="Information", message="คุณตั้งชื่อไฟล์ผิดกรุณาลองใหม่อีกครั้ง")
            self.csv_program.grid_forget()

    def delete_treeview(self):
        for i in self.status_tree.get_children():
            self.status_tree.delete(i)


obj = UvProgram()
obj.userform()