import React from "react";

type MessageListData = {
    messages: Array<string>
}


export default class MessagesList extends React.Component<MessageListData>
{
    render() {
        return <div>
            <p>API Logs:</p>
            {this.props.messages.map((message, i) =>
                <ul>
                    <li>{message}</li>
                </ul>
            )}
        </div>;
    }
}
