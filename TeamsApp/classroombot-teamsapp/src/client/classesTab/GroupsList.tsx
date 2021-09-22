import React, { FunctionComponent } from "react";
import * as moment from 'moment';
import { Button, Header, Form, FormInput, FormButton } from "@fluentui/react-northstar";


import { Channel, ChatMessageMention, DirectoryObject, Event, Group, OnlineMeeting, OnlineMeetingInfo, User } from "@microsoft/microsoft-graph-types";


type ClassListdata = {
    listData: Group[],
    graphToken: string | undefined,
    graphMeetingUser: User,
    log: Function
}
type ClassListProps = {
    selectedGroup: Group | null,
    newMeetingName: string,
    currentMeeting: OnlineMeeting
}

export default class ClassesList extends React.Component<ClassListdata, ClassListProps>
{
    constructor(props) {
        super(props);
        this.setState({ selectedGroup: null });
    }

    handleNewMeetingNameChange(event) {
        this.setState({newMeetingName: event.target.value});
      }

    render() {
        let output;

        if (this.state?.selectedGroup) {
            let header = this.state.selectedGroup.displayName;
            output =
                <Form
                    onSubmit={() => this.startMeeting(this.state.selectedGroup!)} >
                    <Header as="h1" content={header} />

                    <FormInput
                        label="Meeting subject"
                        required
                        value={this.state.newMeetingName} onChange={e => this.handleNewMeetingNameChange(e)} 
                        showSuccessIndicator={false}
                    />

                    <FormButton content="Start New Class" primary />
                    <Button onClick={async () => this.joinLastClass()} secondary 
                        disabled={this.state.currentMeeting !== null}>Join Last Class</Button>
                    <Button onClick={async () => this.cancelCreateMeeting()} secondary>Cancel</Button>
                </Form>;
        }
        else {
            if (this.props.listData.length == 0) {
                output = <div>No groups found</div>;
            }
            else
                output =
                <div>
                    
                    <h3>Your Groups:</h3>
                
                    <table id="meetingListContainer">
                        <thead>
                            <tr>
                                <th>Group name</th>
                            </tr>
                        </thead>
                        <tbody>

                            {this.props.listData.map((group, i) =>
                                <tr className="meetingItem">
                                    <td className="meetingSubject">{group.displayName}</td>
                                    <td>
                                        <Button onClick={async () => this.createMeeting(group)} primary>Start Meeting</Button>
                                    </td>
                                </tr>
                            )}
                        </tbody>
                    </table>
                </div>;
        }


        return <div>{output}</div>;
    }

    hasTeamsMeeting(meeting: Event): boolean {
        return meeting.onlineMeeting !== null;
    }

    createMeeting(group: Group) {
        this.setState({ selectedGroup: group });
    }
    cancelCreateMeeting() {
        this.setState({ selectedGroup: null });
    }

    joinLastClass()
    {
        window.open(this.state.currentMeeting?.joinWebUrl!);
    }

    async startMeeting(group: Group) {
        this.createNewMeeting(group)
            .then(async newMeeting => {

                const channelId = await this.getDefaultChannelId(group);

                this.postMeetingToGroup(group, newMeeting, channelId)
                    .then(async () => {

                        // Join bot
                        await this.joinBotToCall(newMeeting.joinWebUrl!)
                            .then(async () => {

                                this.setState({ currentMeeting: newMeeting });

                                // Everyone to pass through lobby 1st
                                this.props.log("Configuring lobby...");
                                await this.setLobbyBypass(newMeeting, false);

                                this.props.log("All done. Opening meeting in new tab");
                                this.joinLastClass();
                            })
                            .catch(error => {
                                this.props.log('Error from bot API: ' + error);
                            });
                    });


            })
            .catch(error => {
                this.props.log('Error: ' + error);
            });
    }

    
    async getGroupDirectoryObjects(group: Group): Promise<Array<DirectoryObject>> {

        this.props.log("Getting general channel ID...", true);

        // https://docs.microsoft.com/en-us/graph/api/group-list-members
        const endpoint = `https://graph.microsoft.com/v1.0/groups/${group.id}/members`;
        const requestObject = {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                "authorization": "bearer " + this.props.graphToken
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

        this.props.log("Getting user ...", true);

        // https://docs.microsoft.com/en-us/graph/api/user-get
        const endpoint = `https://graph.microsoft.com/v1.0/users/${dirOjbect.id}`;
        const requestObject = {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                "authorization": "bearer " + this.props.graphToken
            }
        };


        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();

        console.info("Got user result");
        console.info(responsePayload);

        return responsePayload;
    }

    async getDefaultChannelId(group: Group): Promise<string> {

        this.props.log("Getting general channel ID...", true);

        const endpoint = `https://graph.microsoft.com/v1.0/teams/${group.id}/primaryChannel`;
        const requestObject = {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                "authorization": "bearer " + this.props.graphToken
            }
        };


        const response = await fetch(endpoint, requestObject);
        const responsePayload = await response.json();

        console.info("Got channel result");
        console.info(responsePayload);


        const channels: Array<Channel> = responsePayload.value;
        return channels[0].id!;
    }

    async createNewMeeting(group: Group): Promise<OnlineMeeting> {

        this.props.log("Creating new meeting...", true);

        // https://docs.microsoft.com/en-us/graph/api/resources/onlinemeeting
        let data: any = {
            "lobbyBypassSettings":
            {
                "scope": "organizer"
            },
            "allowedPresenters": "organizer",
            "subject" : this.state.newMeetingName,
            "participants":
            {
                "organizer": {
                    "identity": { "@odata.type": "#microsoft.graph.identitySet" },
                    "upn": this.props.graphMeetingUser.userPrincipalName,
                    "role": "presenter"
                },
                "attendees":
                    [
                        {
                            "identity": { "@odata.type": "#microsoft.graph.identitySet" },
                            "upn": group.mail,
                            "role": "attendee"
                        }
                    ]
            }
        };

        const endpoint = `https://graph.microsoft.com/v1.0/users/${this.props.graphMeetingUser.id}/onlineMeetings`;
        const requestObject = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                "authorization": "bearer " + this.props.graphToken
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
        this.props.log("Publishing meeting to group channel...", true);

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
                                displayName: user.displayName                            }
                        }
                    });
            });
        }


        let data: any = {
            "body": {
                "contentType": "html",
                "content": `<div>${meeting.subject} - <a href="${meeting.joinWebUrl}">join class now</a></div>
                            <div>${membersHtml}</div>`
            },
            "mentions": mentions
        };

        // https://docs.microsoft.com/en-us/graph/api/channel-post-messages
        const endpoint = `https://graph.microsoft.com/v1.0/teams/${group.id}/channels/${channelId}/messages`;
        const requestObject = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                "authorization": "bearer " + this.props.graphToken
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

    async delay(ms: number) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    async setLobbyBypass(meeting: OnlineMeeting, bypassLobby: boolean) {
        let data: any = {};

        if (bypassLobby) {
            data =
            {
                "lobbyBypassSettings":
                {
                    "scope": "everyone"
                }
            };
        }
        else {
            data =
            {
                "lobbyBypassSettings":
                {
                    "scope": "organizer"
                }
            };
        }

        data.allowedPresenters = "organizer";

        const endpoint = `https://graph.microsoft.com/v1.0//users/${this.props.graphMeetingUser.id}/onlineMeetings/${meeting.id}`;
        const requestObject = {
            method: 'PATCH',
            headers: {
                'Content-Type': 'application/json',
                "authorization": "bearer " + this.props.graphToken
            },
            body: JSON.stringify(data)
        };


        return await fetch(endpoint, requestObject)
            .then(async response => {
                if (response.ok) {

                    const responsePayload = await response.json();

                    console.info("Got meeting update result");
                    console.info(responsePayload);

                    return Promise.resolve();
                }
                else {
                    return Promise.reject(`Got error response ${response.status} from Graph API.`);
                }
            });
    }


    async getMeeting(meetingInfo: OnlineMeetingInfo): Promise<OnlineMeeting> {

        const endpoint = `https://graph.microsoft.com/beta/users/${this.props.graphMeetingUser.id}/onlineMeetings?$filter=JoinWebUrl%20eq%20'${meetingInfo.joinUrl}'`;
        const requestObject = {
            method: 'GET',
            headers: {
                "authorization": "bearer " + this.props.graphToken
            }
        };


        return await fetch(endpoint, requestObject)
            .then(async response => {
                if (response.ok) {

                    const responsePayload = await response.json();

                    console.info("Got meeting details");
                    console.info(responsePayload);

                    return Promise.resolve(responsePayload.value[0]);
                }
                else {
                    return Promise.reject(`Got error response ${response.status} from Graph API.`);
                }
            });
    }

    async joinBotToCall(joinUrl: string) {
        const data =
        {
            "JoinURL": joinUrl,
            "DisplayName": "ClassroomBot"
        };

        const endpoint = `https://${process.env.BOT_HOSTNAME}/joinCall`;
        const requestObject = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
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
