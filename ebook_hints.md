# Walk-through for Windows users

Do you want to use the [Standard Ebooks](https://standardebooks.org/contribute) tools, but you have a Windows computer? 
In this walk-through for Windows users, you will use a USB flash drive to run a Linux-based operating system. 
[Raspberry Pi Desktop](https://www.raspberrypi.com/software/raspberry-pi-desktop/) is a Linux-based operating system
and software bundle that looks like an old version of Windows.
(This approach is an alternative to installing a Linux-based operating system on your computer, 
such as [Windows Subsystem for Linux (WSL)](https://docs.microsoft.com/en-us/windows/wsl/install).)

## Setup steps (done once)

#### Set up Raspberry Pi Desktop on a USB flash drive

1. Insert an empty USB flash drive in your Windows computer. \
‣ The USB flash drive should be 16 MB or larger and USB 3.0 or higher.
(Avoid USB 2, which is too slow for this use.)
If reusing a USB flash drive, check that it really is empty,
without any old files you still want.
1. Download [Raspberry Pi Desktop](https://www.raspberrypi.com/software/raspberry-pi-desktop/).
1. Download [balenaEtcher](https://www.balena.io/etcher/).
1. Open the downloaded file named balenaEtcher...Setup.exe. \
‣ balenaEtcher installs and opens. It's a disk utility needed for the next step.
1. Flash Raspberry Pi Desktop to the USB flash drive with these steps:
   1. In balenaEtcher, click **Flash from file**.
   1. Select the downloaded file named 20...raspios...iso and click **Open**.
   1. Click **Select target**, select the USB flash drive, and click **Select**.
   1. Click **Flash**.
   1. Click **Yes** to the question, "Do you want to allow this app to make changes to your device?"
   1. After the "Flash Completed" message, click the **X** that closes balenaEtcher.
   1. Remove the USB flash drive.
   1. Later, delete the two files you downloaded, 20...raspios...iso and balenaEtcher...Setup.exe.
1. Search online for the steps to boot your computer from a USB flash drive. \
‣ For example, searching Google for **[boot thinkpad from usb](https://www.google.com/search?q=boot+thinkpad+from+usb&oq=boot+thinkpad+from+usb)**
finds [these steps](https://support.lenovo.com/us/en/solutions/ht118361-how-to-boot-from-a-usb-drive-thinkpad), 
"Power on the system.... press F12....
Use the arrow key to...select the USB drive.... press Enter."
   1. (Optional) Write a label and stick it on the USB flash drive. \
‣ For example, the label might say, "R Pi for ebooks. F12 to boot USB. Run w/ persistence."
   1. (Optional) Open this page on your phone to view the next steps while you restart and configure.
1. Boot from the USB flash drive, using the steps for your computer. \
‣ For example, the following steps are for a ThinkPad.
   1. Insert the USB flash drive.
   1. Click the **Windows** icon, select **Power**, and select **Restart**. \
‣ Or when your computer is off, start it.
   1. Press and hold the key that opens boot options, such as **F12**.
   1. Select the USB flash drive on the menu of boot options and press **Enter**.
   1. Select **Run with persistence** and press **Enter**. \
‣ With persistence, your settings and work will be saved.
1. Configure Rapsberry Pi Desktop with these steps:
   1. At the welcome message, click **Next**.
   1. Select your country, language, and time zone and click **Next**.
   1. Choose and type a username and password and click **Next**.
   1. If using Wi-Fi, select your network, click **Next**,
type the Wi-Fi password, and click **Next**.
   1. At the "Update Software" message, click **Next**. \
‣ The updates may take an hour or more.
   1. Click **OK** to the message, "System is up to date."
   1. At the "Setup Complete" message, click **Restart.** \
‣ If your computer doesn't automatically boot from the USB flash drive,
then repeat the steps you used earlier.
1. Continue to configure Raspberry Pi Desktop with these steps:
   1. (Optional) Click the **Web Browser** icon (a globe) and open this page,
if you've been viewing it on your phone. 
   1. (Optional) Move the menu bar to the bottom, like the Windows taskbar, with these steps:
      1. Click the **Raspberry Pi** icon, select **Preferences**, and select **Appearance Settings**.
      1. Click the **Menu Bar** tab, select **Bottom**, and click **OK**.
   1. (Optional) If you're using multiple screens, arrange them with these steps:
      1. Click the **Raspberry Pi** icon, select **Preferences**, and select **Screen Configuration**.
      1. Click and drag the screen icons to position.
      1. To change which screen is primary, click **Layout**,
select **Screens**, select the screen, and select **Primary**. 
      1. Click **Apply**.
      1. Click **No** to the message about rebooting.
      1. Click **Close**.
   1. (Optional) Click the **Bluetooth** icon and select **Turn Off Bluetooth**.
   1. (Optional) Require a user name and password with these steps:
      1. Click the **Raspberry Pi** icon, select **Preferences**, and select **Raspberry Pi Configuration**.
      1. On the System tab, click to turn off Auto login and click **OK**.
   1. (Optional) Click the **Raspberry Pi** icon, select **Logout**, and select **Reboot**. 
      1. Type your password and press **Enter**.
      1. Click the **Web Browser** icon (a globe) and open this page again.

#### Set up the Standard Ebooks toolset on the USB flash drive

1. In the web browser, sign in to [Github](https://github.com/login).
   1. If you don't have a Github account, [sign up](https://github.com/signups).
   1. Click **Save** to the message from the Password Manager.
1. Install the font used by Standard Ebooks with these steps:
   1. In the web browser, open [League Spartan](https://www.theleagueofmoveabletype.com/league-spartan), click **Download**, and close the page.
   1. Click the **File Manager** icon (two folders).
   1. Click the **Downloads** folder.
   1. Right-click the LeagueSpartan...zip file and select **Extract to**.
   1. Type `.fonts` as the location to extract to, and click **Extract**.
   1. Right-click the LeagueSpartan...zip file, select **Move to Trash**, and click **Yes**.
   1. Click the **X** that closes the File Manager.

To be continued . . .
