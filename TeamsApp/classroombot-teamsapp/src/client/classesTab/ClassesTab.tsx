import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";
import { useState, useEffect, useCallback } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import MessagesList from './MessagesList';
import GroupsList from './GroupsList';
import { Group, User } from "@microsoft/microsoft-graph-types";

/**
 * Implementation of the Classes content page
 */
export const ClassesTab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [consentUrl, setConsentUrl] = useState<string>();
    const [user, setUser] = useState<User>();
    const [error, setError] = useState<string>();
    const [nextPageUrl, setNextPageUrl] = useState<string | null>();
    const [allGroups, setAllGroups] = useState<Group[]>([]);
    const [messages, setMessages] = useState<Array<string>>();

    const [ssoToken, setSsoToken] = useState<string>();
    const [msGraphOboToken, setMsGraphOboToken] = useState<string>();

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };

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
        }
    }, [inTeams]);

    const loadUserData = useCallback(async () => {
        if (!msGraphOboToken) { return; }

        // Load groups user is in - https://docs.microsoft.com/en-us/graph/api/group-list?view=graph-rest-1.0
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

        getGroups();


    }, [msGraphOboToken]);

    const getGroupsForUrl = useCallback(async url => {
        if (!msGraphOboToken) { return; }

        const requestObject = {
            method: 'GET',
            headers: {
                "authorization": "bearer " + msGraphOboToken
            }
        };

        await fetch(url, requestObject)
            .then(async response => {
                if (response.ok) {

                    const responsePayload = await response.json();
                    const nextPageUrl : string = responsePayload["@odata.nextLink"];
                    if (nextPageUrl)
                    {
                        setNextPageUrl(nextPageUrl);
                    }
                    else 
                        setNextPageUrl(null);

                    setAllGroups(oldGroups => oldGroups.concat(responsePayload.value));
                }
                else {
                    alert(`Got response ${response.status} from Graph. Check permissions?`);
                }
            })
            .catch(error => {
                alert('Error loading from Graph: ' + error.error?.response?.data?.error);
            });


    }, [msGraphOboToken]);

    const getGroups = useCallback(async () => {
        if (!msGraphOboToken) { return; }

        // Use beta endpoint so we can filter on groups with Teams only
        const endpoint = `https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$select=id,displayName`;
        
        await getGroupsForUrl(endpoint);

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


    const [ignored, forceUpdate] = React.useReducer(x => x + 1, 0);

    const logMessage = ((log: string, clearPrevious: boolean) => {
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
                        
                        {messages &&
                            <MessagesList messages={messages} />
                        }
                        <div>
                            {allGroups &&
                                <div>
                                    <GroupsList listData={allGroups} graphToken={msGraphOboToken} graphMeetingUser={user!} log={logMessage} />
                                </div>
                            }
                        </div>
                        {error &&
                            <div>
                                <div><Text content={`An SSO error occurred ${error}`} /></div>
                                {error == 'consent_required' ?
                                    <div>
                                        <p>You need to grant this application the right permissions to your data.</p>
                                        <a href={consentUrl} target="_blank">Grant access (new window)</a>
                                    </div>
                                    : null}
                            </div>
                        }
                        {nextPageUrl &&
                            <p><Button content="Load next page" tinted onClick={() => getGroupsForUrl(nextPageUrl)}></Button></p>
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
