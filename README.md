# k
SharePoint helper functions for doing common admin tasks from the console

### Usage

k.js is written as a userscript, so you can include it as a file for autoloading in Tampermonkey or Greasemonkey.  
k.js logs a "k" to the console to indicate it's been loaded.  
All commands are in the k namespace, e.g. 

```
await k.fetch("/web/allproperties")
```
lists out all the web properties.

<a href="http://www.wtfpl.net/"><img
       src="http://www.wtfpl.net/wp-content/uploads/2012/12/wtfpl-badge-4.png"
       width="80" height="15" alt="WTFPL" /></a>
