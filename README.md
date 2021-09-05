# eIRe - Electronic Infrared Enhancement

This application allows use of almost any IR remote to control applications on Windows.  IR button presses are converted to keyboard inputs, configurable on an application by application basis, or globally.  It is also possible to execute system commands, for example to launch an application.

Hardware is needed to receive the IR signals from the IR remore.  eIRe uses the [Arduino IRMP SimpleReceiver](https://github.com/ukw100/IRMP/tree/master/examples/SimpleReceiver) example and a TSOP2238 receiver module.  See the [IRMP schematics](https://github.com/ukw100/IRMP) for details.  This will output the decoded IR signals as serial data.  eIRe listens on a COM port for this data.  When data is received, eIRe uses a lookup table to find the appropriate keystrokes to send to the application in focus on Windows.
