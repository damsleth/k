// ==UserScript==
// @name         k365 - the SharePoint Online console helper library
// @namespace    k
// @version      2.0
// @description  do hard task easily in sharepoint online
// @author       @damsleth
// @match        https://*.sharepoint.com/*
// @grant        none
// ==/UserScript==


(function() {

    window.k = { }

    k.switchSharePointExperience = async () => await cookieStore.get('splnu').then(d=>d.value=='0'?cookieStore.set('splnu','1'):cookieStore.set('splnu','0')).then(location.reload())

    // logs and reloads
    k.logAndReload = (str, lvl = "log") => {
        console[lvl](str);
        window.setTimeout(() => location.reload(), 500)
    }

    // fetch (some _api endpoint)
    k.fetch = async (endPoint) => {
        let response = await fetch(`${_spPageContextInfo.siteAbsoluteUrl}/_api/${endPoint}`, {
            credentials: "include", headers: { accept: "application/json;odata=verbose" }
        })
        if (response.ok) {
            let r = await response.json()
            return r ? r.d ? r.d.d ? r.d.d : r.d : r : null
        } else return null
    }

    // get a bearer token you can use for whatever
    k.getToken = async() => {
        let token = await fetch('/_api/contextinfo',{method:"POST",headers:{credentials:"include",accept:"application/json;odata=nometadata"}})
        .then(d=>d.json().then(r=>r.FormDigestValue))
        return token
    }

    // Change web properties - eg, k.ChangeWebProp("Title", "LolSite")
    k.ChangeWebProp = (prop, value) => fetch(`${_spPageContextInfo.webAbsoluteUrl}/_api/web`, {
        method: "MERGE", credentials: "include", headers: {
            "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
            "Content-type": "application/json;odata=verbose",
        }, body: `{"__metadata":{"type": "SP.Web"},"${prop}":"${value}"}`
    })
        .then(res => res.ok ? k.logAndReload(`Successfully set property \n "${prop}": "${value}"`) : k.logAndReload(`Failed setting "${prop}"`, "err"))
        .then(window.setTimeout(() => location.reload(), 500))



    // Sets the value of a property on a sitefield - eg. the "indexed" property of "Title" - k.SetFieldPropValue("Title","indexed","true")
    k.SetFieldPropValue = (field, prop, value) => fetch(`${_spPageContextInfo.webAbsoluteUrl}/_api/web/fields/GetByInternalNameOrTitle('${field}')`, {
        method: "MERGE",
        credentials: "include",
        headers: {
            "X-RequestDigest": document.getElementById("__REQUESTDIGEST").value,
            "Content-type": "application/json;odata=verbose",
        },
        body: `{"__metadata":{"type": "SP.Field"},"${prop}":"${value}"}`
    }).then(res => res.ok ? k.logAndReload(`Successfully set "${prop}" for "${field}" to "${value}"`) : k.logAndReload(`Failed setting "${prop}" on "${field}"`, "err"))

    console.log('%c K ', 'background: #5555ff; color: #eee');
})();
