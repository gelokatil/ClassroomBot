import React, { FunctionComponent } from "react";
import * as moment from 'moment';
import { Button, LeftParenthesisKey } from "@fluentui/react-northstar";

import { Event, OnlineMeeting, OnlineMeetingInfo, User } from "@microsoft/microsoft-graph-types";


type ClassListdata = {
    listData: Event[],
    graphToken: string | undefined,
    graphMeetingUser: User,
    log: Function
}


export default class ClassesList extends React.Component<ClassListdata>
{
    render() {
        let output;
        if (this.props.listData.length == 0) {
            output = <div>No events found in your calendar today.</div>;
        }
        else
            output =
                <table id="meetingListContainer">
                    <thead>
                        <tr>
                            <th>Subject</th>
                            <th>Start</th>
                            <th>End</th>
                        </tr>
                    </thead>
                    <tbody>

                        {this.props.listData.map((meeting, i) =>
                            <tr className="meetingItem">
                                <td className="meetingSubject">{meeting.subject}</td>
                                <td className="meetingDate">{(moment(meeting.start?.dateTime)).format('DD-MMM-YYYY HH:mm:ss')}</td>
                                <td className="meetingDate">{(moment(meeting.end?.dateTime)).format('DD-MMM-YYYY HH:mm:ss')}</td>
                                <td>
                                    <Button onClick={async () => await this.startMeeting(meeting)} disabled={!this.hasTeamsMeeting(meeting)}>Start Meeting</Button>
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>;

        return <div>{output}</div>;
    }

    hasTeamsMeeting(meeting: Event) : boolean {
        return meeting.onlineMeeting !== null;
    }

    async startMeeting(meeting: Event) {

        let url = meeting.onlineMeeting?.joinUrl ? meeting.onlineMeeting?.joinUrl : "";

        var meetingInstance = await this.getMeeting(meeting.onlineMeeting!);

        this.props.log("Configuring meeting...", true);

        // Allow bot to join directly
        await this.setLobbyBypass(meetingInstance, true);

        // Join bot
        await this.joinBotToCall(url)
            .then(async () => {
                
                // Everyone to pass through lobby 1st
                this.props.log("Configuring lobby...");
                await this.setLobbyBypass(meetingInstance, false);

                this.props.log("All done. Opening meeting in new tab");
                window.open(url);

            })
            .catch(error => {
                this.props.log('Error from bot API: ' + error);
            });
    }

    async setLobbyBypass(meeting: OnlineMeeting, bypassLobby: boolean) {
        let data = {};

        if (bypassLobby) {
            data =
            {
                "lobbyBypassSettings":
                {
                    "scope": "everyone"
                }
            };
        }
        else
        {
            data =
            {
                "lobbyBypassSettings":
                {
                    "scope": "organizer"
                }
            };
        }

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

    
    async getMeeting(meetingInfo: OnlineMeetingInfo) : Promise<OnlineMeeting> {

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
