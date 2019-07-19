Office.initialize = function () {
}

function defaultStatus(event) {
    /** HELPER VARIABLES AND FUNCTIONS **/
    // Reading a JSON file
    function readJSONFile(file, callback) {
        var rawFile = new XMLHttpRequest();
        rawFile.overrideMimeType("application/json");
        rawFile.open("GET", file, true);
        rawFile.onreadystatechange = function () {
            if (rawFile.readyState === 4 && rawFile.status == "200")
                //Sets callback to contain the acronym objects
                callback(rawFile.responseText);
        }
        rawFile.send(null);
    }

    // Looks through email body for a matching acronyms and replaces the words span elements.
    function replace(body, acronyms) {
        var index = [];
        var name = [];
        // Checking each acronym
        for (var i = 0; i < acronyms.length; i++) {
            var last = -1;
            //var regex = new RegExp('\\b' + acronyms[i].name + '\\b');
            var loc = body.search('\\b' + acronyms[i].name + '\\b');
            // How many times is it found
            while (loc > last) {
                index[loc] = '<span class="acronym" title="' + acronyms[i].def + '"><u>' + acronyms[i].name + '</u></span>'; //mark acronymss
                name[loc] = acronyms[i].name;
                last = loc;
                loc = body.search('\\b' + acronyms[i].name + '\\b');
            }
        }
        // Going backward through the index
        for (var i = index.length; i > 0; i--) {
            // Replace if something exists at that location
            if (index[i] != null) {
                body = body.substring(0, i) + index[i] + body.substring(i + name[i].length);
            }
        }
        return body;
    }

    // Helper function to add a status message to the info bar.
    function statusUpdate(icon, text) {
        Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
            type: "informationalMessage",
            icon: icon,
            message: text,
            persistent: false
        });
    }

    /** BEGINNING OF SCRIPT **/
    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            var emailBody;
            Office.context.mailbox.item.body.getAsync("html", function callback(result) {
                //Contains the email body
                emailBody = result.value;
                var bodyLoc = emailBody.indexOf("<body");
                emailBody = emailBody.substring(bodyLoc);
                readJSONFile("data.json", function (acronyms) {
                    //Replaces words in email body based on acronyms found in the JSON file
                    emailBody = replace(emailBody, JSON.parse(acronyms));
                    console.log(emailBody);
                    Office.context.mailbox.item.body.setAsync(emailBody, { coercionType: "html" }, function callback(response) {
                        statusUpdate("icon16", "Anti-Acronym plugin " + response.status);
                    });
                });
            });

        });
    });
    /** END OF SCRIPT **/
} 