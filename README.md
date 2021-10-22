# Outlook Security Tamer

Removes safelinks and 'external sender' warnings. Based on a few hypothesis:
 * I get a lot of external emails. Seeing the banner all the time habituates me to ignore it. (See todos too)
 * The banner ruins message preview. The one line preview in Outlook and Outlook for Android just shows the warning. If this warning was integrated into outlook properly it would be much less annoying.Perhaps where the 'high importance' or 'external images have been blocked' information is display.
 * Safelinks ruin text only emails, and make it very hard to inspect the actual destination of URLs.
 * They allow for comprehensive tracking of link clicking by Microsoft.
 * The security benefits are questionable, given that browsers have become quite resilient and do checks themselves (e.g. [Google Safe Browsing](https://en.wikipedia.org/wiki/Google_Safe_Browsing)).

# Installation
Either compile this yourself with [Visual Studio](https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-outlook?view=vs-2019), or use the click-once installer. Works on Windows only (see below).

# Todos
 * Could keep the external email banner for sender email addresses seen for the first time, or those that we haven't send email to ourselves yet. 
 * Could keep safelinks for some domains, or provide a way to whitelist certain domains.
 * Apparently the click-once installer is not very good. I have seen other projects use [Wix](https://github.com/7coil/DiscordForOffice), or perhaps [InnoSetup](https://github.com/bovender/VstoAddinInstaller).
 * Some sort of build pipeline?
 
 # Outlook on Mac OS

 Outlook on Mac OS does not support VB COM / .net add-ins. They only support js / typescript add-ins, which have rather limited abilities - in particular they can't rewrite incoming emails. It might be easiest to implement this desired functionality using a short tool based on [exchangelib](https://github.com/ecederstrand/exchangelib). This is a popular package and perhaps a tool that can be borrowed to implement the desired functionalities exists already ([thumbscr-ews](https://github.com/sensepost/thumbscr-ews), [ubersicht-widget-Exchange](https://github.com/rtphokie/ExchangeMeetings.widget), )