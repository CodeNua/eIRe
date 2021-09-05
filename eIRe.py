# -*- coding: utf-8 -*-
"""
Created on Wed Apr 14 13:10:46 2021

@author: codeNua

eIRe
Copyright (C) 2021  CodeNua

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

import os, time
import psutil, win32process, win32gui
import keyboard
import serial
import subprocess

# GUI Related Imports
from infi.systray import SysTrayIcon    # For System Tray Icon
import ctypes                           # For message box (Mbox)
import signal                           # For keyboard interrupts

# https://github.com/Infinidat/infi.systray/blob/develop/README.md

# --------   Settings   --------
configFilename = "irremote.ini"     # Config file from which IR commands-to-keystrokes are loaded

keyMap = { "pgup":"page up",        # Some irconfig files might use names different to 'keyboard' library
           "pgdn":"page down" }

# -------- End Settings --------

Protocol = 00
Address  = 00
Command  = 00
activeWindowProcess = "FirstRun" 

# Get the name of the active window process
def active_window_process_name():
    pid = win32process.GetWindowThreadProcessId(win32gui.GetForegroundWindow())     #Produces list of PIDs of active windows
    return(psutil.Process(pid[-1]).name())                                          #pid[-1] Return process name of active window

# Open the IR Port
def open_IR_port(IR_Comm_Port, IR_Baudrate):
    ser = serial.Serial(
        port=IR_Comm_Port,\
        baudrate=IR_Baudrate,\
        parity=serial.PARITY_NONE,\
        stopbits=serial.STOPBITS_ONE,\
        bytesize=serial.EIGHTBITS,\
            timeout=0.01)
    
    return(ser)

# Load the IR commands-to-keystrokes and application names 
def load_config_file():
    
    irConfig = {}
    
    with open(os.path.join(os.getcwd(), configFilename), 'r') as f: # open in readonly mode
            
        # Assume that a line is normally "key=value"
        # Lines starting '[' are "Sections"
        # Lines starting ';' are comments
        
        Section = "null"
        
        for line in f:
            
            if line[0] == "[":      # Section
                
                Section = line.rstrip("\r\n")               # Remove any trailing new lines
            
                irConfig[Section] = {}
                #print(line)
            
            elif line[0] ==";":    # Comment
                # Do nothing
                pass
            else:                   # key=value
                
                x = line.split("=")
                key = x[0].strip()                          # Trim whitepace
                try:
                    value = x[1]
                    value = value.split(";")[0].strip()     # Remove comments, trim whitespace
                except:
                    value = ""                              # Some keys have blank values
            
                irConfig[Section][key] = value
                
    return(irConfig)

# Look up the dictionaries for an IR (Address, Command, Application) and return the command
def IR_to_keyCommand(Address, Command, Application):
    
    print("Input:\t\t Address: %s \t Command: %s \t Application: %s" % (Address, Command, Application))
    
    try:
        AddressName = irConfig["[HandsetName]"][str(Address)]
    
        print("Address: \t", AddressName )
    
        try:
            CommandName = irConfig[str("[" + AddressName + "]")][str(Command)]
            
            print("Command: \t", CommandName )
            
            try:
                KeyCommand = irConfig[str("[" + Application + "]")][str(CommandName)]
            except:
                KeyCommand = irConfig[str("[Default]")][str(CommandName)]
                
                print("KeyCmmd:\t", KeyCommand )
        
        except:
            KeyCommand = ""
    except:
        KeyCommand = ""
   
    return(KeyCommand)

# Determine if the keyCommand is a keystroke or something else (e.g. system command 'run')
def keyCommand_interpret(keyCommandIn):
    
    if keyCommandIn[0:4] == "{run":
        x = keyCommandIn[4:]
        x = x.strip("(\")}")
        sysCommand = x
        print(sysCommand)           
        subprocess.Popen(sysCommand)        # Run system command
        
        
    else:
        KeyStroke = keyCommandIn.replace("}","+").replace("{","").rstrip("+")
        
        try:
            KeyStroke = keyMap[KeyStroke]
        except:
            pass
        
        print("KeyStrk:\t", KeyStroke)
        
        if len(KeyStroke) > 0:
            keyboard.send(KeyStroke)        # Press and Release keystroke


##########################        
# ------  SysTray  -------
def Mbox(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

def systray_on_quit_callback(systray):
    # Stop the main loop
    global runLoop
    print("Stopping the IR Loop.")
    runLoop = False
    
def systray_check_status(systray):
    
    global irCount, Protocol, Address, Command, activeWindowProcess, IR_Comm_Port
    
    text = "Comm Port:        \t\t" + str(IR_Comm_Port) + "\n" + \
           "IR Input Count:   \t\t" + str(irCount) + "\n" + \
           "\n" + \
           "Last IR Input Received" + "\n" + \
           "Protocol:         \t\t" + str(Protocol) + "\n" + \
           "Address:          \t\t" + str(Address) + "\n" + \
           "Command:          \t\t" + str(Command) + "\n" + \
           "Application:      \t\t" + str(activeWindowProcess)
                      
    Mbox("IR Status", text, 0)
    
def systray_help(systray):
    text =  "Settings are loaded from irremote.ini."
    Mbox("IR Help", text, 0)
    
# Catches CTRL+C keyboard interrupt to exit		   
def keyboardInt_handler(signal, frame):
    print('A keyboard interrupt occured.')
    systray.shutdown()

    


##########################        
# --------  MAIN  --------

# Setup keyboard interrupt handler
signal.signal(signal.SIGINT, keyboardInt_handler)

# Load the IR Config file
irConfig = load_config_file()            

# Open IR Serial Port
IR_Comm_Port = irConfig["[Config]"]["IR_Comm_Port"]
IR_Baudrate  = int(irConfig["[Config]"]["IR_Baudrate"])

try:
    ser = open_IR_port(IR_Comm_Port, IR_Baudrate)
except:
    try:
        print("Comm Port wouldn't open.  Trying to close and open again.")
        ser.close()
        ser = open_IR_port(IR_Comm_Port, IR_Baudrate)
    except:
        print("Comm Port wouldn't open.  Is the port set correctly?")
        
print("IR Port Opened on:       ", IR_Comm_Port, "-", IR_Baudrate, "baud")



# Setup System Tray Icon

sysTrayTitle = "IR Remote on " + IR_Comm_Port
sysTrayIcon  = "icon.ico"          
menu_options = (("Check Status", None, systray_check_status), \
               ("Help", None, systray_help),)                           # Text, Icon, Command

systray = SysTrayIcon(sysTrayIcon, sysTrayTitle, menu_options, on_quit=systray_on_quit_callback)

systray.start()


serialIn = ""

# Set parameters for Repeat Delay (i.e. button held down)
# Objective is to make it repeat like a "real keyboard"
repeatDelay         = int(irConfig["[Config]"]["repeatDelay"])  # Delay time in ms
repeatCount         = 0
repeatTimeStart     = time.time()
previousCommand     = 0
previousTime        = time.time()

irCount = 0

runLoop = True   # Set FALSE for Debug, TRUE for run

while runLoop:

    serialIn = ser.readline().decode("utf-8") 
        
    if len(serialIn) > 19 and serialIn[0] == "P":
        #print(str(len(serialIn)) + "\t\t" + str(serialIn))
        #print(serialIn)
        
        commandTime = time.time()
        
        # # This is for debug - Stops the loop of a specific IR key
        # # "P=RC5 A=0x1E C=0x29" is the blue button on Hauppauge remote
        # # In normal operation, leave commented out
        # if serialIn[0:19] == "P=RC5 A=0x1E C=0x29":
        #     print("Stopping the loop!")
        #     runLoop = False
        
        x =  serialIn.split(" ")    
        
        Protocol = x[0][2:]
        Address = int(x[1][2:], 0)
        Command = int(x[2][2:], 0)
        try:
            Repeat = x[3][0]
        except:
            Repeat = ""
        
        elapsedTime = int( (commandTime - previousTime  ) * 1000 )
        
        
        if Command == previousCommand:  
            # Possible Repeat
        
            if elapsedTime < 125:                  
                # Short - Assume Key Held Down (125 ms seems to work well)
                
                print(">> Elapsed Time: ", elapsedTime)
                if repeatCount == 0:
                    repeatTimeStart = commandTime
                    print(">> Repeat Time:", repeatTimeStart)
                repeatCount += 1
                repeatOkay = 0
                
            else:                                   
                # Long - Assume Key Released and pressed again as separate event
                
                repeatCount = 0
                repeatOkay = 1
            
            repeatTime = int( (commandTime - repeatTimeStart)*1000 ) 
            
            if repeatTime > repeatDelay:
                # Repeat Delay - After timeout, send keys anyway
                repeatOkay = 1
        else:                           # New Command
            repeatTimeStart = commandTime
            repeatOkay = 1

        
        # If all okay, truck on and do the commands       
        if repeatOkay == 1:
            
            # Print the command and some debug stuff
            print("InfoBlast: ", Protocol, Address, Command, Repeat, repeatCount, repeatOkay, int( (commandTime - previousTime)*1000 ))
            # serialIn = ""
            
            # Get the name of the active (in-focus) window            
            activeWindowProcess = active_window_process_name()[:-4]  # Active Process Name without '.exe' extension
            
            # Lookup the config data to find out what to do with the IR command
            stringLookup = IR_to_keyCommand(Address, Command, activeWindowProcess)
            
            # Make it so!
            keyCommand_interpret(stringLookup)
            
            
        previousCommand = Command
        previousTime = commandTime
        
        irCount += 1

print("Closing serial port.")
ser.close()

print("Shutdown!")



        
