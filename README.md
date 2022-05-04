# k
SharePoint helper functions for doing common admin tasks from the console.  
`k.js` is meant for on premises (i.e. SharePoint 2013, 2016 and 2019)  
`k365.js` is for SharePoint Online (i.e. Microsoft 365, *yourtenant*.sharepoint.com )


### Usage

k.js is written as a userscript, so you can include it as a file for autoloading in Tampermonkey or Greasemonkey.  
k.js logs a "k" to the console to indicate it's been loaded.  
All commands are in the k namespace, e.g. 

```
await k.fetch("/web/allproperties")
```
lists out all the web properties.
```
await k.getToken()
```
fetches a bearer token from `/_api/contextinfo` you can use for whatever
```
await k.ChangeWebProp("Title","My new site title")
```
Changes the title of the site you're on

### Requirements
Chromium or Firefox browser  
Tampermonkey or Greasemonkey  
A SharePoint Site  

<a href="http://www.wtfpl.net/"><img
       src="http://www.wtfpl.net/wp-content/uploads/2012/12/wtfpl-badge-4.png"
       width="80" height="15" alt="WTFPL" /></a>
