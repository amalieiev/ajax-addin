Office.initialize = function () {};

const messages = [];

function debugMessage(text) {
    messages.push(text);

    const messageHTML = `
		<table>
			${messages
                .map((message) => {
                    return `<tr><td>${message}</td></tr>`;
                })
                .join("")}
		</table>
	`;

    return new Promise((resolve) => {
        Office.context.mailbox.item?.body.setSignatureAsync(
            messageHTML,
            { coercionType: Office.CoercionType.Html },
            () => {
                resolve();
            }
        );
    });
}

async function onMessageComposeHandler(event) {
    try {
        debugMessage("Start");
    } catch (error) {
        debugMessage("Error");
    }

    debugMessage("End");

    event.completed();
}

Office.actions.assotiate("OnNewMessageCompose", onMessageComposeHandler);
