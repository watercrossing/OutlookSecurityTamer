# Outlook Security Tamer

Removes safelinks and 'external sender' warnings. Based on a few hypothesis:
 * I get a lot of external emails. Seeing the banner all the time habituates me to ignore it. (See todos too)
 * The banner ruins message preview. The one line preview in Outlook and Outlook for Android just shows the warning.
 * Safelinks ruin text only emails, and make it very hard to inspect the actual destination of URLs.
 * They allow for comprehensive tracking of link clicking by Microsoft.
 * The security benefits are questionable, given that browsers have become quite resilient and do checks themselves.

# Installation
Either compile this yourself with [Visual Studio](https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-outlook?view=vs-2019), or use the click-once installer.

# Todos
 * Could keep the external email banner for sender email addresses seen for the first time.
 * Could keep safelinks for some domains, or provide a way to whitelist certain domains.
 * Apparently the click-once installer is not very good. I have seen other projects use [Wix](https://github.com/7coil/DiscordForOffice), or perhaps [InnoSetup](https://github.com/bovender/VstoAddinInstaller).
 * Some sort of build pipeline?
 
 