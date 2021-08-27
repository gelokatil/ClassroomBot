import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect, useCallback } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import ClassesList from './ClassesList';

/**
 * Implementation of the Classes content page
 */
export const ClassesTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [recentEvents, setRecentEvents] = useState<any[]>();

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
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

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
                    console.info(responsePayload);
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
        getTodaysMeetings();
    }, [msGraphOboToken]);

    const exchangeSsoTokenForOboToken = useCallback(async () => {
        const response = await fetch(`/exchangeSsoTokenForOboToken/?ssoToken=${ssoToken}`);
        const responsePayload = await response.json();
        if (response.ok) {
            setMsGraphOboToken(responsePayload.access_token);
        } else {
            if (responsePayload!.error === "consent_required") {
                setError("consent_required");
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
                                    <ClassesList listData={recentEvents} />
                                    <Button onClick={() => getTodaysMeetings()}>Refresh</Button>
                                </div>
                            }
                        </div>
                        {error &&
                            <div>
                                <div><Text content={`An SSO error occurred ${error}`} /></div>
                                {error == 'consent_required' ? <Text>Grant access</Text> : null}
                            </div>
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
