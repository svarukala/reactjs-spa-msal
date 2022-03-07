import React, { useState, useEffect } from "react";
import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from "@azure/msal-react";
import { Pivot, PivotItem } from 'office-ui-fabric-react';
import { loginRequest, oboRequest } from "./authConfig";
import { PageLayout } from "./components/PageLayout";
import { ProfileData } from "./components/ProfileData";
import { callMsGraph } from "./graph";
import Button from "react-bootstrap/Button";
import "./styles/App.css";
import SPOReusable from "./components/SPOReusable";
import MSGReusable from "./components/MSGReusable";

/**
 * Renders information about the signed-in user or a button to retrieve data about the user
 */

const ProfileContent = () => {
    const { instance, accounts } = useMsal();
    const [graphData, setGraphData] = useState(null);
    const [ssoToken, setSsoToken] = useState(null);
    const [accessToken, setAccessToken] = useState(null);

    useEffect(() => {    

      }, []);     

    function RequestProfileData() {
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance.acquireTokenSilent({
            ...loginRequest,
            account: accounts[0]
        }).then((response) => {
            callMsGraph(response.accessToken).then(response => setGraphData(response));
        });
    }
    function RequestTokenForOBOApp() {
        // Silently acquires an access token which is then attached to a request for MS Graph data
        instance.acquireTokenSilent({
            ...oboRequest,
            account: accounts[0]
        }).then((response) => {
            setSsoToken(response.idToken);
            setAccessToken(response.accessToken);
        });
    }
    return (
        <>
            <h5 className="card-title">Welcome {accounts[0].name}</h5>
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
                    <PivotItem headerText="SPO REST API">
                        <SPOReusable idToken={accessToken} />
                    </PivotItem>
                    <PivotItem headerText="MS Graph REST API">
                        <MSGReusable idToken={accessToken} />
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
                <ProfileContent />
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
