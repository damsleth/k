// ==UserScript==
// @name         k - the SharePoint console helper library
// @namespace    k
// @version      0.1
// @description  do hard task easily in sharepoint
// @author       @damsleth
// @match        http://*/*
// @grant        none
// ==/UserScript==


(function() {

    window.k = { }

    // logs and reloads
        k.logAndReload = (str, lvl = "log") => {
            console[lvl](str);
            window.setTimeout(() => location.reload(), 500)
        }

        k.fetch = async (endPoint) => {
        let response = await fetch(`${_spPageContextInfo.siteAbsoluteUrl}/_api/${endPoint}`, {
            credentials: "include", headers: { accept: "application/json;odata=verbose" }
        })
        if (response.ok) {
            let r = await response.json()
            return r ? r.d ? r.d.d ? r.d.d : r.d : r : null
        } else { return null }
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


    // Adds a scriptlink customaction with the given name (key) and url - e.g. k.AddCustomAction("k.lol.js", "/siteassets/js/k.lol.js")
    k.AddCustomAction = (name,scriptblock,sequence) => {
        let seq = sequence;
        if(!sequence){ seq = Math.floor(Math.random() * 1000) };
        let ctx = SP.ClientContext.get_current();
        let uc = ctx.get_site().get_userCustomActions();
        let uca = uc.add();
        uca.set_location("ScriptLink");
        uca.set_sequence(seq);
        uca.set_title(name);
        uca.set_name(name)
        uca.set_description(`Adds ${name}`);
        uca.set_scriptBlock(`${scriptblock}`);
        uca.update();
        ctx.executeQueryAsync(() => k.logAndReload(`Added CustomAction ${name}`), () => k.logAndReload(`Failed adding customaction ${name}`, "err"));
    }

    // Adds a SOD scriptlink customaction with the given name (key) and url - e.g. k.AddCustomAction("k.lol.js", "/siteassets/js/k.lol.js")
    k.AddCustomActionSOD = (name, url) => {
        let seq = Math.floor(Math.random() * 1000)
        let ctx = SP.ClientContext.get_current();
        let uc = ctx.get_site().get_userCustomActions();
        let uca = uc.add();
        uca.set_location("ScriptLink");
        uca.set_sequence(seq);
        uca.set_title(name);
        uca.set_name(name);
        uca.set_description(`Adds ${name}`);
        uca.set_scriptBlock(`SP.SOD.registerSod('${name}', '${url}?rev=' + new Date().getDay());LoadSodByKey('${name}');`);
        uca.update();
        ctx.executeQueryAsync(() => k.logAndReload(`Added CustomAction ${name}`), () => k.logAndReload(`Failed adding customaction ${name}`, "err"));
    }

    // Deletes a custom action by name. If none is specified, lists all customactions
    k.DeleteCustomAction = (customActionName) => {
        let ucaID;
        const ctx = SP.ClientContext.get_current();
        let ucaColl = ctx.get_site().get_userCustomActions();
        ctx.load(ucaColl);
        ctx.executeQueryAsync(() => {
            let en = ucaColl.getEnumerator();
            if (!customActionName) { console.log(`No User Custom Action Specified, listing all UCAs in site`) }
            while (en.moveNext()) {
                let currentUCA = en.get_current();
                let ucaName = currentUCA.get_name();
                if (!customActionName) { console.log(ucaName); }
                if (ucaName === customActionName) {
                    ucaID = currentUCA.get_id();
                }
            }
            if (ucaID) {
                let ucaToDelete = ucaColl.getById(ucaID);
                ucaToDelete.deleteObject();
                ctx.load(ucaToDelete);
                ctx.executeQueryAsync(() => {
                    let ucaToDeleteName = ucaToDelete.get_name();
                    k.logAndReload(`Custom action ${ucaToDeleteName} deleted`);
                }, () => k.logAndReload(`Unable to delete customaction ${ucaToDelete.get_name()}`,"err"));
            } else {
                if (customActionName) {
                    k.logAndReload(`User custom action ${customActionName} not found`, "warn");
                }
            }
        });
    }


    // Enables a web feature by guid
    k.AddWebFeature = (guid) => {
        const ctx = SP.ClientContext.get_current()
        let web = ctx.get_web()
        ctx.load(web)
        let feats = web.get_features()
        let featGuid = new SP.Guid(`{${guid}}`)
        feats.add(featGuid, true, SP.FeatureDefinitionScope.web)
        ctx.load(feats)
        ctx.executeQueryAsync(() => k.logAndReload("enabled suitenav,maybe?"))
    }


    // Sets the JSLink on a given field - e.g. k.SetJsLinkOnField("Title","/siteassets/jslink/jsLolLink.js")
    k.SetJsLinkOnField = (fieldName, jsLink) => {
        const JSLinkPrefix = "clienttemplates.js|";
        const ctx = new SP.ClientContext.get_current();
        let field = ctx.get_web().get_availableFields().getByInternalNameOrTitle(fieldName);
        ctx.load(field);
        ctx.executeQueryAsync(() => {
            field.set_jsLink(`${JSLinkPrefix}${jsLink}`);
            field.updateAndPushChanges(true);
            ctx.executeQueryAsync(() => {
                k.logAndReload(`Set JSLink with url ${jsLink} for field ${fieldName}.\n${field.get_jsLink()}`);
            });
        });
    }


    // Returns the given jslink for a sitefield
    k.GetJSLinkForField = (fieldName) => {
        const ctx = SP.ClientContext.get_current();
        let field = ctx.get_web().get_availableFields().getByInternalNameOrTitle(fieldName);
        ctx.load(field);
        ctx.executeQueryAsync(() => {
            console.log(field.get_jsLink());
        });
    }
    console.log("%c k ","background:#000")
    })();
