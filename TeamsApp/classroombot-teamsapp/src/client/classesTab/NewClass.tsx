import React from "react";
import { Button, Header, Input } from "@fluentui/react-northstar";
import { Channel, ChatMessageMention, DirectoryObject, Event, Group, OnlineMeeting, User } from "@microsoft/microsoft-graph-types";

type NewClassProps = {
    graphToken: string | undefined,
    graphMeetingUser: User,
    log: Function,
    cancelNewMeeting: Function,
    selectedGroup : Group
}
type NewClassState = {
    newMeetingName: string,
    currentMeeting: OnlineMeeting | null,
    loading: boolean
}

export default class NewClass extends React.Component<NewClassProps, NewClassState> {
    constructor(props) {
        super(props);
        this.state = { newMeetingName: "", currentMeeting: null, loading: false };
    }

    handleNewMeetingNameChange(event) {
        this.setState({ newMeetingName: event.target.value });
    }

    render() {
        const title : string = `New Online Meeting in '${this.props.selectedGroup?.displayName}'`;
        return <div>
            <Header as="h3" content={title} />
            <p>Enter meeting details below.</p>
            <Input
                label="Meeting subject"
                required
                value={this.state.newMeetingName} onChange={e => this.handleNewMeetingNameChange(e)}
            />
            <div>
                <Button primary onClick={() => this.startMeeting(this.props.selectedGroup!)} 
                    disabled={this.state.loading}>Start New Class</Button>
                    {this.state.currentMeeting &&
                        <Button onClick={async () => this.joinLastClass()} secondary>Join Created Meeting</Button>
                    }
                <Button onClick={async () => this.props.cancelNewMeeting()} secondary>Cancel</Button>
            </div>
        </div>;
    }

    hasTeamsMeeting(meeting: Event): boolean {
        return meeting.onlineMeeting !== null;
    }

    joinLastClass() {
        window.open(this.state.currentMeeting?.joinWebUrl!);
    }

    async startMeeting(group: Group) {

        // Create new meeting with form values
        this.createNewMeeting(group)
            .then(async newMeeting => {

                // Post meeting details
                const channelId = await this.getDefaultChannelId(group);
                this.postMeetingToGroup(group, newMeeting, channelId)
                    .then(async () => {
                        
                        // Remember meeting in state
                        this.setState({ currentMeeting: newMeeting });

                        // Join bot
                        await this.joinBotToCall(newMeeting.joinWebUrl!)
                            .then(async () => {


                                this.props.log("All done. Opening meeting in new tab");
                                this.joinLastClass();
                            })
                            .catch(error => {
                                this.props.log(`Error from ${process.env.BOT_HOSTNAME} API: ${error}`);
                            });
                    });
            });
    }

    async getGroupDirectoryObjects(group: Group): Promise<Array<DirectoryObject>> {

        this.props.log("Getting group members...");

        // https://docs.microsoft.com/en-us/graph/api/group-list-members
        const endpoint = `https://graph.microsoft.com/v1.0/groups/${group.id}/members`;
        const requestObject = {
            method: "GET",
            headers: {
                method: "POST",
                authorization: "bearer " + this.props.graphToken
            }
        };

        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();

        console.info("Got group-members result");
        console.info(responsePayload);

        const members: Array<DirectoryObject> = responsePayload.value;
        return members;
    }

    async getUser(dirOjbect: DirectoryObject): Promise<User> {

        // https://docs.microsoft.com/en-us/graph/api/user-get
        const endpoint = `https://graph.microsoft.com/v1.0/users/${dirOjbect.id}`;
        const requestObject = {
            method: "GET",
            headers: {
                method: "POST",
                authorization: "bearer " + this.props.graphToken
            }
        };

        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();

        console.info("Got user result");
        console.info(responsePayload);

        return responsePayload;
    }

    async getDefaultChannelId(group: Group): Promise<string> {

        this.props.log("Getting default channel ID...");

        // https://docs.microsoft.com/en-us/graph/api/team-get-primarychannel
        const endpoint = `https://graph.microsoft.com/v1.0/teams/${group.id}/primaryChannel`;
        const requestObject = {
            method: "GET",
            headers: {
                authorization: "bearer " + this.props.graphToken
            }
        };

        const response = await fetch(endpoint, requestObject);
        const responsePayload : Channel = await response.json();

        console.info("Got channel result");
        console.info(responsePayload);

        return responsePayload.id!;
    }

    async createNewMeeting(group: Group): Promise<OnlineMeeting> {

        this.props.log("Creating new online meeting...", true);

        // https://docs.microsoft.com/en-us/graph/api/resources/onlinemeeting
        const data: any = {
            lobbyBypassSettings:
            {
                scope: "organizer"
            },
            allowedPresenters: "organizer",
            subject: this.state.newMeetingName,
            participants:
            {
                organizer: {
                    identity: { "@odata.type": "#microsoft.graph.identitySet" },
                    upn: this.props.graphMeetingUser.userPrincipalName,
                    role: "presenter"
                },
                attendees:
                    [
                        {
                            identity: { "@odata.type": "#microsoft.graph.identitySet" },
                            upn: group.mail,
                            role: "attendee"
                        }
                    ]
            }
        };

        const endpoint = `https://graph.microsoft.com/v1.0/users/${this.props.graphMeetingUser.id}/onlineMeetings`;
        const requestObject = {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                authorization: "bearer " + this.props.graphToken
            },
            body: JSON.stringify(data)
        };

        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();
        console.info("Got meeting create result");
        console.info(responsePayload);
        return responsePayload as OnlineMeeting;
    }

    async postMeetingToGroup(group: Group, meeting: OnlineMeeting, channelId: string) {

        const groupDirObjects = await this.getGroupDirectoryObjects(group);
        this.props.log("Publishing meeting to group channel...");

        let membersHtml : string = '';
        let mentions : Array<ChatMessageMention> = [];
        if(groupDirObjects)
        {
            let userQueries : Array<Promise<User>> = [];
            groupDirObjects.map((member, i) =>
            {
                userQueries.push(this.getUser(member));
            }
            );

            let users : Array<User> = [];
            const allUserQs = await Promise.all(userQueries);
            allUserQs.map(user=> { 
                console.log(user);
                users.push(user);
            }
            );

            users.map((user, i) => 
            { 
                membersHtml += `<at id="${i}">${user.displayName}</at>, `;
                mentions.push(
                    {
                        id: i,
                        mentionText: user.displayName,
                        mentioned:
                        {
                            user:
                            {
                                id: user.id,
                                displayName: user.displayName                            
                            }
                        }
                    });
            });
        }

        let data: any = {
            body: {
                contentType: "html",
                content: `<div>${meeting.subject} - <a href="${meeting.joinWebUrl}">join class now</a></div>
                            <div>${membersHtml}</div>`
            },
            mentions: mentions
        };

        // https://docs.microsoft.com/en-us/graph/api/channel-post-messages
        const endpoint = `https://graph.microsoft.com/v1.0/teams/${group.id}/channels/${channelId}/messages`;
        const requestObject = {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                authorization: "bearer " + this.props.graphToken
            },
            body: JSON.stringify(data)
        };

        const response = await fetch(endpoint, requestObject);
        if (response.ok) {

            const responsePayload = await response.json();
            console.info("Got meeting create result");
            console.info(responsePayload);
            return responsePayload as Event;
        }
        else {
            return Promise.reject(`Got response ${response.status} from Graph. Check permissions?`);
        }
    }

    async joinBotToCall(joinUrl: string) {
        const data =
        {
            JoinURL: joinUrl,
            DisplayName: "ClassroomBot"
        };

        // Call our own bot URL to have it join our meeting
        const endpoint = `https://${process.env.BOT_HOSTNAME}/joinCall`;
        const requestObject = {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(data)
        };

        await fetch(endpoint, requestObject)
            .then(async response => {
                if (response.ok) {

                    const responsePayload = await response.json();

                    this.props.log("Bot has accepted join request");
                    console.info("Got bot join response");
                    console.info(responsePayload);

                    return Promise.resolve(responsePayload);
                }
                else {
                    return Promise.reject(`Got error response ${response.status} from Bot API.`);
                }
            });
    }
}
