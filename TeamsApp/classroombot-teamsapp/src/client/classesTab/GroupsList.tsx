import React from "react";
import { Button } from "@fluentui/react-northstar";

import { Event, Group, User } from "@microsoft/microsoft-graph-types";

type ClassListdata = {
    listData: Group[],
    graphToken: string | undefined,
    graphMeetingUser: User,
    log: Function,
    newMeeting: Function
}

export default class ClassesList extends React.Component<ClassListdata>
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

        if (this.props.listData.length === 0) {
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
                                    <Button onClick={async () => this.props.newMeeting(group)} primary>Start Meeting</Button>
                                </td>
                            </tr>
                        )}
                    </tbody>
                </table>
            </div>;
        


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
}
