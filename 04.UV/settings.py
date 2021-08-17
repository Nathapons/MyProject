import serial
import serial.tools.list_ports as tool
import time


class ReadProgram:
   def get_ports(self):
      # Function for getting port number
      ports = tool.comports()
      uv_port = "No Connect"
      for port, desc, hwid in sorted(ports):
         if desc == 'Prolific USB-to-Serial Comm Port':
            uv_port = port
            break
      return uv_port

   def read_value(self, event, sec):
      pass

# port = "COM11"
# baud = 19200
 
# ser = serial.Serial(port, baud)
 
# if ser.isOpen():
#      print(ser.name + ' is open...')
 
# cmd = input("Enter command or 'exit':")
 
# while cmd != "exit":
#    ser.write(cmd.encode()+b'\r\n')
#    out = ser.readline()
#    print('Receiving...'+ out.decode())
 
# ser.close()
# exit()