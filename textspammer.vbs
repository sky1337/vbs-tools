set a = createobject("wscript.shell")

mytxt=inputbox("Send what text?","Type some text","Type Here")
mynum=inputbox("How many times to send the text?","spamnumber","50") 
myspeed=inputbox("How fast to Type..In milisecs!","delay","100") 
mywait=inputbox("Wait Before Sending","Wait?","5")

msgbox("You have " & mywait & " secs to put focus on your target text area!")
wscript.sleep (mywait*1000) 
for i=1 to mynum 		'count down from mynum variable
	a.sendkeys (mytxt)       'Sends the text you typed in the mytxt prompt
	a.sendkeys ("{ENTER}")   'presses the enter key to send your text you may change it to the apropriate key that sends the msg in your chat
	wscript.sleep (myspeed)   'sleeps OR waits the amount of Milseconds you typed in the Mywait prompt
next
msgbox("Spam Finished")
