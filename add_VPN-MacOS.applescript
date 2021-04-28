#user inputs email + and set to variable we can use later
set user_input to display dialog "What is your work email?" default answer ""
set user_id to text returned of user_input

#user inputs password and set to variable we can use later -> **plaintext**
set user_input2 to display dialog "What is your rippling password?" default answer ""
set user_pw to text returned of user_input2

#user inputs PSK + and set to variable we can use later
set user_input3 to display dialog "What is the pre-shared key?" default answer ""
set vpn_psk to text returned of user_input3

#our initialized variables with values set
set vpn_name to "PT-VPN"
set vpn_url to "pt-brwqkwngtbc.dynamic-m.com"


# start apple interactive scripting - aka formatting going to shit ######
#########################################################################

#closes any existing network settings window; which will break the script
tell application "System Preferences"
	reveal pane "Network"
	quit
	delay 1
end tell


#symphony of chaos insues, 
tell application "System Preferences"
	reveal pane "Network"
	activate
	tell application "System Events"
		tell process "System Preferences"
			tell window 1
				click button "Add Service"
				tell sheet 1
					click pop up button 1
					click menu item "VPN" of menu 1 of pop up button 1
					delay 1
					click pop up button 2
					click menu item "L2TP over IPSec" of menu 1 of pop up button 2
					keystroke tab
					keystroke vpn_name
					click button "Create"
				end tell
				delay 0.5
				keystroke tab
				delay 1
				keystroke tab
				keystroke vpn_url
				delay 0.5
				keystroke tab
				keystroke user_id
				delay 0.5
				tell group 1
					click button 1
					keystroke user_pw
					delay 0.5
					keystroke tab
					keystroke vpn_psk
					delay 0.5
				end tell
				tell sheet 1
					click button "OK"
				end tell
				click button "Apply"
			end tell
		end tell
	end tell
end tell
