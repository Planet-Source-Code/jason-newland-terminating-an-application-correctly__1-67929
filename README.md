<div align="center">

## Terminating an application correctly


</div>

### Description

This is to explain the all too common problem for beginners to VB of "Why does my application hang when I close it and why is the compiled EXE version still in Windows Processes?".
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason Newland](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-newland.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-newland-terminating-an-application-correctly__1-67929/archive/master.zip)





### Source Code

```
When beginning in Visual Basic 6, its a new experience with RAD (Rapid Application Development) and you have absolutely no idea how the execution structure works.
In this short article, I'm going to discuss the all-too common problem of application termination doesn't terminate at all, but rather the IDE (Intergrated Development Environment) hangs and clicking on END just crashes it, or the compiled EXE stays in Windows Processes and doesn't terminate.
First, let's explain what your program actually does when you click on the "X" close button (you can also apply this principle to ANY VB form).
When you close a form, a set of EVENTS (or conditional sections) FIRE or execute. The first event fired is the form's QueryUnload event. In QueryUnload is where you would prompt the user if they really want to close the form (program) or not and this is where you could CANCEL the whole unload process. If no user intervention is detected in QueryUnload, unloading continues and it is then passed on to the forms Unload event.
The Unload event is fired when the form has actually unloaded all together, not the misconception that it IS unloading. The idea here, would be NOT to put any code in this event at all, unless you are checking that the form has unloaded for a particular function call reason.
If a form is SUBCLASSED (a window hooking call for example to maybe "skin" the form) do not UNSUBCLASS (unhook the window) during the Unload event. This is usually the very reason the application won't terminate properly some of (if not all of) the time.
After the form has successfully unloaded, the Terminate event fires. This, in the main form terms, means that the application has successfully unloaded all resources, forms and objects from memory and is now ready to be executed again another time.
OK, so lets assume you have a subclassed form with a skin, and in the Load event of the form, you've passed the following call "Call SubClass(Me.hwnd)". On termination of the program, you want to pass this call "Call UnSubClass(Me.hwnd)", to unsubclass it again.
What you would do here is make a Private Sub called, say "DoUnload" at the bottom of the forms code and put the command in there, together with Unload Me (and all other Set object = Nothing calls). You would call this sub from the forms QueryUnload event and by rights it should terminate correctly from there on.
What if it doesn't? Subclassing can be a pain in the rear at the best of times. If you still experience unloading problems, try placing a TIMER on the form, in the DoUnload sub, remove the Unload Me, set the timer's interval property to say, 200 milliseconds, then set its enabled property to TRUE. In the timer's timer sub, set the enabled property back to FALSE the call Unload Me.
At no point, should you ever use the END command. This command terminates the application at the point it is called and, for instance, if called withing the forms Unload event, does not allow the form to correctly terminate or fire it's Terminate event. Thus, this will lead to it being stuck in processes of Windows, crashing the IDE, hanging the EXE and causing your CPU usage to sky-rocket to 100%, amoung other ridiculous, hair-pulling, stress related "why is it doing that?!" problems. Never use it, unless you are first loading the program and the user doesn't have a DLL or something installed, wrong OS version, etc, or you are doing error handling that an error has occured, popup a message box stating the error, then end the program.
I hope this article has shed some light on the way VB handles form termination. Any question, simply put them in the comments below :). Have a nice day and have fun!
```

