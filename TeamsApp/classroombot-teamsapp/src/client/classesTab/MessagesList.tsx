import React from "react";

type MessageListData = {
    messages: Array<string>
}


export default class MessagesList extends React.Component<MessageListData>
{
    render() {


        return <div>
            <p>Logs:</p>
            {this.props.messages.map((message, i) =>
                <div>
                    <p>{message}</p>

                </div>
            )}
        </div>;
    }
}
