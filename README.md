# iaer-js
Intune, autopilot &amp; entra (device) remover, now in JavaScript!

### What is this? 😕
A web application for fetching device data, selecting devices and deleting them from Intune, Autopilot and Entra.</br></br>
The latest version I implemented of this was in PowerShell. The modules, the multiple scripts and the errors...</br>
Well, (hopefully) not anymore! I've tested this with a few devices and have received successful results!</br></br>
As of now, I have a few checks, cleanup and more failchecks to implement, use with caution.

### How / when does it work? 🤔
On the web application users can submit serials to the input area. This works best, when the device's name includes the serial
since Entra devices cannot be fetched via serials.</br></br>
Fetched devices will be listed with latest login details, IDs from the different systems, serials, model info and more.</br>
Each successfully listed device may be selected and deleted with the press of a button. The process is automatic and if</br>
the deletion fails at any step, the program interrupts itself for analysation.</br></br>
Failchecks have been implemented in many places. Deletion does not happen, unless the user desides to so. Fetching doesn't delete anything.

This application works with an app registration to Entra. The following API permissions might be required:</br>
 > AuditLog.Read.All</br>
 > Device.Read.All</br>
 > Device.ReadWrite.All</br>
 > DeviceManagementManagedDevices.Read.All</br>
 > DeviceManagementManagedDevices.ReadWrite.All</br>
 > DeviceManagementServiceConfig.Read.All</br>
 > DeviceManagementServiceConfig.ReadWrite.All</br>
 > User.Read.All</br>

### Installation and usage 🛠️
Clone the repository to your wished location and use ```npm install``` to install required packages.</br>
Add your client and tenant IDs into ```/js/appSettings.js``` along with your client secret</br>
Use ```node main.js``` to start the application, which opens to ```127.0.0.1:4070```.
