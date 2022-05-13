import React, { useState, useEffect } from "react";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { AuthorizationUrlRequest } from "@azure/msal-browser";
import { Pivot, PivotItem, IPivotStyleProps, IPivotStyles } from 'office-ui-fabric-react';
import { graphConfig, loginRequest, oboRequest } from "./authConfig";
import { PageLayout } from "./components/PageLayout";
import { ProfileData } from "./components/ProfileData";
import { callMsGraph } from "./graph";
import Button from "react-bootstrap/Button";
import "./styles/App.css";
import { Providers, ProviderState, SimpleProvider } from '@microsoft/mgt-element';
import SPOReusable from "./components/SPOReusable";
import MSGReusable from "./components/MSGReusable";
import MGTReusable from "./components/MGTReusable";

/**
 * Renders information about the signed-in user or a button to retrieve data about the user
 */

const CollabContent = () => {
    const { instance, accounts } = useMsal();
    const [graphData, setGraphData] = useState(null);
    const [mailData, setMailData] = useState(null);
    const [ssoToken, setSsoToken] = useState(null);
    const [accessToken, setAccessToken] = useState(null);
    const [mgtAccessToken, setMgtAccessToken] = useState(null);
    const [error, setError] = useState(null);

    useEffect(() => {    
        if (!Providers.globalProvider) {
            console.log('Initializing global provider');
            Providers.globalProvider = new SimpleProvider(async ()=>{return getAccessTokenForMGT()});  
            Providers.globalProvider.setState(ProviderState.SignedIn);
        } 
      }, []);     

    function RequestProfileData() {
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        }).then((response) => {
            callMsGraph(response.accessToken, graphConfig.graphMeEndpoint).then(response => setGraphData(response));
        });
    }
    function RequestMailData() {
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        }).then((response) => {
            callMsGraph(response.accessToken, graphConfig.graphMailEndpoint).then(response => setMailData(response));
        }).catch(error => setError(error));
    }
    function RequestTokenForOBOApp() {
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance.acquireTokenSilent({
            ...oboRequest,
            account: accounts[0]
        }).then((response) => {
            setSsoToken(response.idToken);
            setAccessToken(response.accessToken);
        })
        .catch((err) => {
            console.log(err);
            console.log("Silent Failed");
            oboRequest.account = accounts[0];
            instance.acquireTokenPopup(oboRequest).then((tokenResponse) => {
            setSsoToken(tokenResponse.idToken);
            setAccessToken(tokenResponse.accessToken);
          })
          .catch((error) => {
            console.error(error);
            // I haven't implemented redirect but it is fairly easy
            console.error("Maybe it is a popup blocked error. Implement Redirect");
            return null;
          });
        });
    }

    function getAccessTokenForMGT() {
        console.log("Getting access token async");
        if(mgtAccessToken) return mgtAccessToken;
        setCurrentAccount(mgtTokenrequest);
        console.log(currentAccount);
        return msalInstance
            .acquireTokenSilent(mgtTokenrequest)
            .then((tokenResponse) => {
            console.log("Inside Silent");
            console.log("Access token: "+ tokenResponse.accessToken);
            console.log("ID token: "+ tokenResponse.idToken);
            setMgtAccessToken(tokenResponse.accessToken);
            return tokenResponse.accessToken;
            })
            .catch((err) => {
            console.log(err);
            console.log("Silent Failed");
            if (err instanceof InteractionRequiredAuthError) {
                return interactionRequired(mgtTokenrequest);
            } else {
                console.log("Some other error. Inside SSO.");
                //const loginPopupRequest: AuthorizationUrlRequest = mgtTokenrequest as AuthorizationUrlRequest;
                const loginPopupRequest = mgtTokenrequest;
                loginPopupRequest.loginHint = loginName;
                return msalInstance
                .ssoSilent(loginPopupRequest)
                .then((tokenResponse) => {
                    setMgtAccessToken(tokenResponse.accessToken);
                    return tokenResponse.accessToken;
                })
                .catch((ssoerror) => {
                    console.error(ssoerror);
                    console.error("SSO Failed");
                    if (ssoerror) {
                    return interactionRequired(mgtTokenrequest);
                    }
                    return null;
                });
            }
            });
    };
    
    return (
        <>
            <h5 className="card-title">Welcome {accounts[0].name}</h5>
            {error && <div className="error">{error.errorMessage}</div>}

            {graphData ? 
                <ProfileData graphData={graphData} />
                :
                <Button variant="secondary" onClick={RequestProfileData}>Request Profile Information</Button>
            }
            
            <br/><br/>
            <Button variant="secondary" onClick={RequestTokenForOBOApp}>Request Token for OBO App</Button>
            {
                accessToken &&
                <Pivot aria-label="Basic Pivot Example">
                    <PivotItem headerText="SPO REST API" backgroundColor="red" >
                        <SPOReusable idToken={accessToken} />
                    </PivotItem>
                    <PivotItem headerText="MS Graph REST API">
                        <MSGReusable idToken={accessToken} />
                    </PivotItem>
                    <PivotItem headerText="MS Graph Toolkit">
                        <MGTReusable />
                    </PivotItem>                                  
                </Pivot>
            }
        </>
    );
};



/**
 * If a user is authenticated the ProfileContent component above is rendered. Otherwise a message indicating a user is not authenticated is rendered.
 */
const MainContent = () => {

    return (
        <div className="App">
            <AuthenticatedTemplate>
                <CollabContent />
            </AuthenticatedTemplate>
            <UnauthenticatedTemplate>
                <h5 className="card-title">Please sign-in to see your profile information.</h5>
            </UnauthenticatedTemplate>
        </div>
    );
};

export default function App() {
    return (
        <PageLayout>
            <MainContent />
        </PageLayout>
    );
}


/*
            <br/><br/>
            {mailData ? 
                <ProfileData graphData={mailData} />
                :
                <Button variant="secondary" onClick={RequestMailData}>Request Mails </Button>
            }
*/