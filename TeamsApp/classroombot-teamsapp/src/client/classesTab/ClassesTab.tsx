import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect, useCallback } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import MessagesList from './MessagesList';
import ClassesList from './ClassesList';
import { User } from "@microsoft/microsoft-graph-types";

/**
 * Implementation of the Classes content page
 */
export const ClassesTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [consentUrl, setConsentUrl] = useState<string>();
    const [user, setUser] = useState<User>();
    const [error, setError] = useState<string>();
    const [recentEvents, setRecentEvents] = useState<any[]>();
    const [messages, setMessages] = useState<Array<string>>();

    const [ssoToken, setSsoToken] = useState<string>();
    const [msGraphOboToken, setMsGraphOboToken] = useState<string>();

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                    setName(decoded!.name);

                    setSsoToken(token);

                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [process.env.TAB_APP_URI as string]
            });

            // Build consent url
            const c = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?" + 
            `client_id=${process.env.MICROSOFT_APP_ID}` + 
            "&response_type=code" + 
            `&redirect_uri=${window.location}` + 
            "&response_mode=query" + 
            "&scope=" + 
            `${process.env.SSOTAB_APP_SCOPES}`;

            setMessages(new Array<string>());

            setConsentUrl(c);
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    const loadUserData = useCallback(async () => {
        if (!msGraphOboToken) { return; }

        // Load user data
        const endpoint = `https://graph.microsoft.com/v1.0/me/`;
        const requestObject = {
            method: 'GET',
            headers: {
                "authorization": "bearer " + msGraphOboToken
            }
        };

        await fetch(endpoint, requestObject)
            .then(async response => {
                if (response.ok) {

                    const responsePayload = await response.json();

                    console.info("Loaded user data:");
                    console.info(responsePayload);
                    setUser(responsePayload);
                }
                else {
                    alert(`Got response ${response.status} from Graph. Check permissions?`);
                }
            })
            .catch(error => {
                alert('Error loading from Graph: ' + error.error.response.data.error);
            });

        getTodaysMeetings();


    }, [msGraphOboToken]);

    const getTodaysMeetings = useCallback(async () => {
        if (!msGraphOboToken) { return; }

        const now = new Date();
        const tomorrow = new Date();
        tomorrow.setDate(tomorrow.getDate() + 1);
        const endpoint = `https://graph.microsoft.com/v1.0/me/calendarview?startdatetime=${now.toISOString()}&enddatetime=${tomorrow.toISOString()}`;
        const requestObject = {
            method: 'GET',
            headers: {
                "authorization": "bearer " + msGraphOboToken
            }
        };

        await fetch(endpoint, requestObject)
            .then(async response => {
                if (response.ok) {

                    const responsePayload = await response.json();

                    console.info("Found events:");
                    console.info(responsePayload.value);
                    setRecentEvents(responsePayload.value);
                }
                else {
                    alert(`Got response ${response.status} from Graph. Check permissions?`);
                }
            })
            .catch(error => {
                alert('Error loading from Graph: ' + error.error.response.data.error);
            });


    }, [msGraphOboToken]);

    useEffect(() => {
        loadUserData();
    }, [msGraphOboToken]);

    const exchangeSsoTokenForOboToken = useCallback(async () => {
        const response = await fetch(`/exchangeSsoTokenForOboToken/?ssoToken=${ssoToken}`);
        const responsePayload = await response.json();
        if (response.ok) {
            setMsGraphOboToken(responsePayload.access_token);
        } else {
            if (responsePayload!.error === "consent_required") {
                setError(`consent_required`);
            } else {
                setError("unknown SSO error");
            }
        }
    }, [ssoToken]);

    useEffect(() => {
        // if the SSO token is defined...
        if (ssoToken && ssoToken.length > 0) {
            exchangeSsoTokenForOboToken();
        }
    }, [exchangeSsoTokenForOboToken, ssoToken]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
        }
    }, [context]);

    const [ignored, forceUpdate] = React.useReducer(x => x + 1, 0);

    const logMessage = ((log : string, clearPrevious : boolean) => {
        if (clearPrevious !== undefined && clearPrevious === true)
            setMessages(new Array<string>());

        console.log(log);
        messages?.push(log);

        // Force render
        forceUpdate(1);
    });


    return (
        <Provider theme={theme}>
            <Flex fill={true} column styles={{
                padding: ".8rem 0 .8rem .5rem"
            }}>
                <Flex.Item>
                    <Header content="Start a class meeting with the ClassroomBot" />
                </Flex.Item>
                <Flex.Item>
                    <div>
                        <div>
                            {recentEvents &&
                                <div>
                                    <h3>Todays Meetings in Your Calendar:</h3>
                                    <ClassesList listData={recentEvents} graphToken={msGraphOboToken} graphMeetingUser={user!} log={logMessage} />
                                    <Button onClick={() => getTodaysMeetings()}>Refresh</Button>
                                </div>
                            }
                        </div>
                        {error &&
                            <div>
                                <div><Text content={`An SSO error occurred ${error}`} /></div>
                                {error == 'consent_required' ? 
                                    <div>
                                        <p>You need to grant this application permissions to your calendar.</p>
                                        <a href={consentUrl} target="_blank">Grant access (new window)</a>
                                    </div> 
                                : null}
                            </div>
                        }
                        {messages &&
                            <MessagesList messages={messages} />
                        }
                    </div>
                </Flex.Item>
                <Flex.Item styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Text size="smaller" content="(C) Copyright Sam Betts" />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
