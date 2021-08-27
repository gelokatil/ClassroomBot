import React, { FunctionComponent } from "react";
import * as moment from 'moment';
import { Button } from "@fluentui/react-northstar";

import { Event } from "@microsoft/microsoft-graph-types";


type Listdata = {
    listData: Event[]
}


export default class ClassesList extends React.Component<Listdata>
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
                                    <Button onClick={async () => await this.startMeeting(meeting)}>Start Meeting</Button>
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>;

        return <div>{output}</div>;
    }

    async startMeeting(meeting: Event) {
        console.log(meeting.onlineMeeting?.joinUrl);
        alert(`Starting ${meeting.subject} at url ${meeting.onlineMeeting?.joinUrl}`);

        const data = 
        {
            "JoinURL": meeting.onlineMeeting?.joinUrl,
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

        
        let url = meeting.onlineMeeting?.joinUrl ? meeting.onlineMeeting?.joinUrl : "";


        await this.joinBotToCall(url)
            .then(() => 
                {
                    window.open(url);
                })
            .catch(error => 
                {
                    alert('Error loading from bot API: ' + error);
                });

    }

    async joinBotToCall(joinUrl : string)
    {
        const data = 
        {
            "JoinURL": joinUrl,
            "DisplayName": "Bot"
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
            .then(async response => 
                {
                    if (response.ok) {
                        
                        const responsePayload = await response.json();

                        console.info("Got bot join response");
                        console.info(responsePayload);

                        return Promise.resolve(responsePayload);
                    }
                    else
                    {
                        return Promise.reject(`Got error response ${response.status} from Bot API.`);
                    }
                });
    }
}
