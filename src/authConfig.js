/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { LogLevel } from "@azure/msal-browser";

/**
 * Configuration object to be passed to MSAL instance on creation. 
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md 
 */
export const msalConfig = {
    auth: {
        clientId: "ba686da8-8cb8-4e41-9765-056a10dee34c",//msaljs-v2-test  
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "/"
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {	
        loggerOptions: {	
            loggerCallback: (level, message, containsPii) => {	
                if (containsPii) {		
                    return;		
                }		
                switch (level) {		
                    case LogLevel.Error:		
                        console.error(message);		
                        return;		
                    case LogLevel.Info:		
                        console.info(message);		
                        return;		
                    case LogLevel.Verbose:		
                        console.debug(message);		
                        return;		
                    case LogLevel.Warning:		
                        console.warn(message);		
                        return;		
                }	
            }	
        }	
    }
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
    scopes: ["user.read"]
};

export const oboRequest = {
    scopes: ["api://3271e1a1-0da7-476b-b573-e360600674a9/access_as_user"]
    //["api://sridev.ngrok.io/c613e0d1-161d-4ea0-9db4-0f11eeabc2fd/access_as_user"]
};

const mgtTokenrequest = {
    scopes: ["Mail.Read","calendars.read", "user.read", "openid", "profile", "people.read", "user.readbasic.all", "files.read", "files.read.all"],
    //process.env.SPFX_MGT_SCOPES.split(","), 
    //['Mail.Read','calendars.read', 'user.read', 'openid', 'profile', 'people.read', 'user.readbasic.all', 'files.read', 'files.read.all'],
    //account: currentAccount,
};

/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages"
};

export const collabAppConfig = {

SPFX_OBOBROKER_URL: "https://azfun.ngrok.io/api/TeamsOBOHelper",

SPFX_MSG_SEARCHQUERY: "https://graph.microsoft.com/v1.0/sites?search=Contoso",

SPFX_SPO_SEARCHQUERY: "https://m365x229910.sharepoint.com/_api/search/query?querytext=%27*%27&selectproperties=%27Author,Path,Title,Url%27&rowlimit=10"

}