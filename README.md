The Merlin Excel add-in was created to supplement Excel with many frequently-needed features that don't exist in native Excel.  The code for Merlin was largely written and/or "borrowed" from the internet by Lonestarjeepin, but with significant contributions by others along the way (noted in the code where applicable since this was created before GitHub).  I'm hoping by converting this from its previous hosting location (merlinaddin.xyz) to GitHub, that contributions from others can help grow the usefulness of the tool.

<B>Download:</B><br>
<OL>
  <LI value="1">You can download the add-in directly from https://github.com/Lonestarjeepin/Merlin/blob/main/Merlin.xlam, then click the download icon or "Raw".  This is the easiest way to get up and running.  Proceed to Install steps below.</LI>
  <blockquote>NOTE: You can also build the add-in yourself if you are comfortable with the VBA editor.  Simply create the necessary modules/classes in VBA Editor, copy contents of the individual module/class files from GitHub, and paste into VBA editor.  I removed the headers that are created while exporting modules/classes so importing these .bas and .cls files from GitHub won't work.  However, I wanted to publish the raw code for branching and PRs in GitHub which I'll use to update the .xlam.</blockquote>
</OL>

<B>Install:</B><br>
<OL>
  <LI value="2">Once downloaded, move the file to a permanent location of your choosing (e.g. Documents).  DO NOT try to open the .xlam file directly.  This will work, but only until you close Excel.  When you re-open Excel, Merlin won't persist.</LI>
  <LI value="3">Once the file is downloaded, open Excel, go to Options, Add-Ins, then click the Go button next to Manage: Excel add-ins.  Click Browse to find the location where you saved the .xlam file, select the Merlin.xlam file, and click OK.  In Excel, you should now see an Add-ins menu item and Merlin will be present.</LI>
</OL>

<B>Update:</B><br>
  <UL>Click the Update Merlin menu item in Merlin.  Your verison will be compared to the latest release version in GitHub.  If there is a new version, it will be downloaded and installed.  Simply restart Excel to use the new version.</UL>

<B>First Usage:</B><br>
  <UL>If you are new to Merlin, click "Merlin ChangeLog and Help" within the add-in and a new worksheet will be created that lists and explains many of the features of the tool.</UL>
