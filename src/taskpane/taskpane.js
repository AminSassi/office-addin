/* global document, Office */  

const apiUrl = "https://192.168.80.101/api/v4/documents/upload/"; // Use HTTPS  
const apiKey = "1f37cc553da59bf5610d847322a4cbf1102234e0"; // Your actual API key  
const documentTypeId = 14; // Document type ID for Mayan  

Office.onReady(async (info) => {  
    if (info.host === Office.HostType.Outlook) {  
        console.log("Outlook context detected.");  
        document.getElementById("sideload-msg").style.display = "none";  
        document.getElementById("app-body").style.display = "flex";  
        document.getElementById("upload").onclick = uploadAttachments;  

        await displayEmailSubject();  
    } else {  
        console.error("This add-in is not running in Outlook.");  
    }  
});  

async function displayEmailSubject() {  
    const item = Office.context.mailbox.item;  

    // Log item details for debugging  
    console.log("Item details:", item);  

    document.getElementById("item-subject").innerText = "Subject: " + (item.subject || "No Subject");  
}  

async function uploadAttachments() {  
    const item = Office.context.mailbox.item;  

    if (item.itemType !== Office.MailboxEnums.ItemType.Message) {  
        alert("This add-in only works with email messages.");  
        return;  
    }  

    // Verify the item object  
    console.log("Item:", item);  

    item.getAsync(["attachments"], async (result) => {  
        if (result.status === Office.AsyncResultStatus.Succeeded) {  
            const attachments = result.value.attachments;  

            if (attachments.length === 0) {  
                alert("No attachments found.");  
                return;  
            }  

            const uploadPromises = attachments.map(async (attachment) => {  
                if (attachment.isInline) return; // Skip inline images  

                try {  
                    const attachmentContent = await getAttachmentContent(attachment.id);  
                    console.log("Uploading:", attachment.name);  
                    await uploadToMayan(attachmentContent, attachment.name);  
                } catch (error) {  
                    console.error(`Error uploading ${attachment.name}:`, error);  
                    alert(`Error uploading ${attachment.name}`);  
                }  
            });  

            await Promise.all(uploadPromises);  
            alert("All uploads complete!");  
        } else {  
            console.error("Error retrieving attachments:", result.error);  
            alert("Error retrieving attachments");  
        }  
    });  
}  

function getAttachmentContent(attachmentId) {  
    return new Promise((resolve, reject) => {  
        Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {  
            if (result.status === Office.AsyncResultStatus.Succeeded && result.value.content) {  
                resolve(result.value.content);  
            } else {  
                reject(result.error || "Failed to get attachment content.");  
            }  
        });  
    });  
}  

async function uploadToMayan(fileContent, fileName) {  
    const formData = new FormData();  
    const blob = new Blob([fileContent], { type: "application/octet-stream" }); // Adjust based on file type  
    formData.append("file", blob, fileName);  
    formData.append("document_type_id", documentTypeId);  

    try {  
        const response = await fetch(apiUrl, {  
            method: "POST",  
            headers: {  
                "Authorization": `Token ${apiKey}`  
            },  
            body: formData  
        });  

        if (response.ok) {  
            console.log(`Successfully uploaded: ${fileName}`);  
        } else {  
            console.error(`Failed to upload: ${fileName} - ${response.statusText}`);  
            alert(`Failed to upload: ${fileName}`);  
        }  
    } catch (error) {  
        console.error(`Error uploading ${fileName}: ${error}`);  
        alert(`Error uploading ${fileName}`);  
    }  
}