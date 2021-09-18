import serial
import serial.tools.list_ports as tool
from tkinter import messagebox as msb


class ReadProgram:
   def get_ports(self):
      # Function for getting port number
      ports = tool.comports()
      uv_port = "No Connect"
      for port, desc, hwid in sorted(ports):
         if 'Prolific USB-to-Serial Comm Port' in desc:
            uv_port = port
            break
      return uv_port

   def read_value(self):
      uv_value = 0
      try:
         port = self.get_ports()

         if port == "No Connect":
            msb.showwarning(
               title='Warning'
               , message='กรุณาตรวจสอบ Port อีกครั้งหนึ่ง'
            )

         ser = serial.Serial(port, 19200)
         cmd = ":03D1"

         while cmd != "exit":
            # time.sleep(0.5)
            ser.write(cmd.encode() + b'\n')
            out = ser.readline()
            value = str(out.decode())
            if '00000000' not in str(out):
               break
         
         uv_value = round(int(value[5:])/10000, 3)
         ser.close()
         return uv_value

      except Exception as err:
         str_error = str(err)
      finally:
         return uv_value