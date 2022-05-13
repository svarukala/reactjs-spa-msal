import React from "react";

/**
 * Renders information about the user obtained from MS Graph
 * @param props 
 */
export const JWTDecodeQuickLink = (props) => {
    console.log(props.token);

    return (
        <div>
            <a href={`https://jwt.ms/#access_token=${props.token}`} target="_blank">{}</a>
        </div>
    );
};