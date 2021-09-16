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
    debugMessage("Start");

    try {
        try {
            debugMessage("Ajax Start");
            await $.ajax({
                url: "https://amalieievfunctions.azurewebsites.net/api/get-signatures",
                dataType: "json",
                headers: { Authorization: "Bearer qwe123" },
            });
            debugMessage("Ajax Success");
        } catch (error) {
            debugMessage("Ajax Error");
        }
    } catch (error) {
        debugMessage("Error");
    }

    debugMessage("End");

    event.completed();
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
