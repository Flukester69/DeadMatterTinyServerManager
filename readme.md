Dead Matter Tiny Server Manager

What's the purpose of this program?

Answer: I wanted a tiny program to monitor my Dead Matter Server without need of installing any 3rd-party software/bloat on the server. This program uses vbscript which can be easily run from Windows just double clicking.

What does it do?

Answer: Simply launches and manages the server and when the server crashes, it will restart the server. The server will be scanned every minute (configurable) to make sure it is up. This script does not work based on memory as the server has been seen to go from 7GB or ram to 10GB easily, but at the same rate it can also go back down to 7GB. When the server crashes it will detect the process is not running anymore and re-kick the server quickly. The manager will report how many times the server has crashed and provide a timestamp of the last restart time. It's a simple utility so there are only 3 buttons to it. Yes - to start, No - to stop and Cancel, to simply exit.

What are the Requirements for usage?

Answer: Configure your game.ini file with the desired configuration. Refer to the DeadMatter channels for more info. Then Copy this file to where the deadmatterServer.exe, modify the vbs file to point to where your Dead Matter server is installed. Then launch! That's it!
