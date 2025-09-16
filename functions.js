function restId(id) {
    return Office.context.mailbox.convertToRestId(
        id,
        Office.MailboxEnums.RestVersion.v2_0
    );
}

function owaLinkFromRestId(rid) {
    return "https://outlook.office.com/mail/deeplink/read/" + encodeURIComponent(rid);
}

function copy(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
        return navigator.clipboard.writeText(text);
    }
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
    return Promise.resolve();
}

function copyOwaLink(event) {
    try {
        const id = Office.context.mailbox.item.itemId;
        const rid = restId(id);
        const link = owaLinkFromRestId(rid);
        copy(link).finally(() => event.completed());
    } catch {
        event.completed();
    }
}

if (typeof Office !== "undefined") {
    Office.initialize = () => {
    };
    Office.onReady(() => {
        Office.actions.associate("copyOwaLink", copyOwaLink);
        window.copyOwaLink = copyOwaLink;
    });
}