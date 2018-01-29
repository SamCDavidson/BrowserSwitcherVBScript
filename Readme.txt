BrowserSwitcherVBScript
Fast lightweight Windows CScript for seamless self executed toggle selection switching of an unprivileged  users default browser choice. Specifically designed for non local administrator accounts with locked down Windows user rights.

In current form, it temporarily adjusts your current IE Default Browser settings too Google Chrome and remains resident until you're finished in Google. Upon confirmation of an OK Dialog Box, your default browser selecion is reset to IE before the VB exits.

Uses VB SendKey instructions along with two second sleep functions, this universal desk top macro simplifies a normally arcane Windows user setting.

A sequentially ordered script, rearranging the orders of execution reverses this script for environments where Chrome is the default browser and IE must be temporarily set for completing some web app functions.

WshShell.Run "%systemroot%\system32\control.exe /name Microsoft.DefaultPrograms /page pageDefaultProgram\pageAdvancedSettings?pszAppName=Internet%20Explorer"
WshShell.Run "%systemroot%\system32\control.exe /name Microsoft.DefaultPrograms /page pageDefaultProgram\pageAdvancedSettings?pszAppName=Google%20Chrome"

Useful inside internal networks where Microsoft Internet Explorer is the preferred desktop browser for reasons such as application compatibility.

When varying applications work better in Google Chrome, this script can be loaded in web page code upon browser agent detection such as automatically from a redirected landing page prompting the end user to accept Microsoft browser warnings required for direct execution.

Four VBMODAL stay on top dialog boxes advises the user their default browser choices is going to be reconfigured for Chrome after clicking OK, then advises the user their default browser choice is now Chrome and another VBMODAL window remains on top while the user seamlessly clicks URL's for effortless opening inside Chrome.

After satisfactory task completion independently determined by the end user, pressing OK reverses the process reopening the Control Panel Default Programs choosing Microsoft Internet Explorer as the default choice of browsers for opening linked URL's.

Narrow in its function, created out of necessity for an organization heavily reliant on MS IE for mission critical web based applications for the indefinite future while phasing in HTML5 applications requiring Chrome. This simple toolkit script preserves an automated end user experience while avoiding porting applications and the associated duplicate server presence.

The organization would've been forced to expel tremendous development costs to port their newly installed HTML5 apps to legacy code since the option of teaching the broad user base the steps to open Control Panel and Default Programs was wholly rejected along with rejecting remote view / RDP remote browser configuration since field office employees would be over slow links or offline completely.
