import serial
import time
 
port = "COM7"
baud = 19200
 
ser = serial.Serial(port, baud)
 
if ser.isOpen():
     print(ser.name + ' is open...')
 
cmd = ":03D1"
 
count = 0
while cmd != "exit":
    ser.write(cmd.encode() + b'\n')
    out = ser.readline()
    value = str(out.decode())
 
    if '00000000' not in str(out):
        count += 1
 
    if count == 10:
        break
 
uv_value = round(int(value[5:])/10000, 3)
print(uv_value)
ser.close()
exit()