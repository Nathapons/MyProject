from tkinter import *
from tkinter import messagebox as msb
from tkinter import ttk
from settings import ReadProgram
import serial.tools.list_ports as tool
import serial


class UvProgram:
    def __init__(self):
        self.setting = ReadProgram()
        self.program()

    def program(self):
        root = Tk()
        
        WIDTH = 600
        HEIGHT  = 500
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        left = int((screen_height/2) - (HEIGHT/2))
        top = int((screen_width/2) - (WIDTH/2))

        # UI Properties
        root.geometry(f'{WIDTH}x{HEIGHT}+{top}+{left}')
        root.title('UV Program')
        root.resizable(0, 0)

        font_style1 = ('Arial', 28)
        font_style2 = ('Arial', 20)

        # Widget
        big_frame = Frame(root)
        big_frame.pack()
        topic = Label(big_frame, text='UV Program', font=font_style1, justify='center')
        topic.grid(row=0, column=0, pady=5)

        self.sec = StringVar()
        select_frame = Frame(big_frame)
        sec_label = Label(select_frame, text='วินาที:', font=font_style2, justify='center')
        sec_spin = Spinbox(select_frame, textvariable=self.sec, from_=1, to=60, font=font_style2, justify='center', width=5)
        select_frame.grid(row=1, column=0)
        sec_label.grid(row=0, column=0, padx=10)
        sec_spin.grid(row=0, column=1, padx=10)

        self.run_program = Button(big_frame, text="Run Program", font=font_style2, justify='center')
        self.run_program.bind('<Button-1>', self.overview)
        self.run_program.grid(row=2, column=0)

        treeview_frame = Frame(big_frame)
        headers = ['Filename', 'Item Code', 'FETL lot', 'Report status']
        self.pull_force_treeview = ttk.Treeview(treeview_frame, column=headers, show='headings',
                                                height=self.treeview_height)
        style = ttk.Style()
        style.configure("Treeview.Heading", font=('Arial', 16, 'bold'))
        style.configure("Treeview", font=('Arial', 12))
        vertical_scrollbar = ttk.Scrollbar(treeview_frame, orient="vertical", command=self.pull_force_treeview.yview)
        self.pull_force_treeview.configure(yscrollcommand=vertical_scrollbar.set)
        for header in headers:
            self.pull_force_treeview.heading(header, text=header)
            if header == 'Filename':
                column_width = 360
            else:
                column_width = 160
            self.pull_force_treeview.column(header, anchor='center', width=column_width, minwidth=0)
            self.pull_force_treeview.bind("<Double-1>", self.pull_force_link_tree)
        root.mainloop()

    def get_ports(self):
          # Function for getting port number
        ports = tool.comports()
        uv_port = "Not Connection"
        for port, desc, hwid in sorted(ports):
            if desc == 'Prolific USB-to-Serial Comm Port':
                uv_port = port
                break
        return uv_port

    def overview(self, event):
        uv_port = self.get_ports()
        if uv_port == "Not Connection":
            msb.showwarning(title="Alarm to user", message="กรุณาตรวจสอบ USB เชื่อมกับ ORC")
        else:
            if self.run_program['text'] == 'Run Program':
                self.run_program['text'] = 'Stop Program'
                self.read_value(uv_port)
            else:
                self.run_program['text'] = 'Run Program'

    def read_value(self, uv_port):
        port = uv_port
        baud = 19200
        
        ser = serial.Serial(port, baud)
        
        if ser.isOpen():
             print(ser.name + ' is open...')
        
        cmd = input("Enter command or 'exit':")
        
        while cmd != "exit":
           ser.write(cmd.encode()+b'\r\n')
           out = ser.readline()
           print('Receiving...'+ out.decode())
        
        ser.close()
        exit()


if __name__ == '__main__':
    app = UvProgram()