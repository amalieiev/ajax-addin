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
            debugMessage("Native fetch Start");
            await fetch(
                "https://amalieievfunctions.azurewebsites.net/api/get-signatures",
                {
                    mode: "cors",
                    method: "POST",
                    headers: {
                        "Content-Type": "application/json",
                        Authorization: "Bearer qwe123",
                    },
                    body: JSON.stringify({
                        sender: "artem.malieiev@gmail.com",
                    }),
                }
            );
            debugMessage("Native fetch Success");
        } catch (error) {
            debugMessage("Native fetch Error");
            debugMessage(error.message ? error.message : error);
        }

        try {
            debugMessage("jQuery.ajax Start");
            await $.ajax({
                url: "https://amalieievfunctions.azurewebsites.net/api/get-signatures",
                dataType: "json",
                headers: { Authorization: "Bearer qwe123" },
            });
            debugMessage("jQuery.ajax Success");
        } catch (error) {
            debugMessage("jQuery.ajax Error");
            debugMessage(error.message ? error.message : error);
        }

        try {
            debugMessage("Text/Plain Start");
            await fetch(
                "https://amalieievfunctions.azurewebsites.net/api/get-signatures",
                {
                    mode: "cors",
                    method: "POST",
                    headers: {
                        "Content-Type": "text/plain",
                    },
                    body: JSON.stringify({
                        sender: "artem.malieiev@gmail.com",
                    }),
                }
            );
            debugMessage("Text/Plain Success");
        } catch (error) {
            debugMessage("Text/Plain Error");
            debugMessage(error.message ? error.message : error);
        }
    } catch (error) {
        debugMessage("Error");
    }

    debugMessage("End");

    event.completed();
}

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
