'use strict';

(function () {
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
    function find(body, acronyms) {
        var found = [];
        // Checking each acronym
        for (var i = 0; i < acronyms.length; i++) {
            //var regex = new RegExp('\\b' + acronyms[i].name + '\\b');
            var loc = body.search('\\b' + acronyms[i].name + '\\b');
            if (loc > -1) {
                var name = "<p><b>" + acronyms[i].name + "</b>";
                var def = acronyms[i].def.substring(acronyms[i].name.length) + "</p>";
                found.push(name + def);
            }
        }
        return found;
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
                readJSONFile("Functions/data.json", function (acronyms) {
                    //Replaces words in email body based on acronyms found in the JSON file
                    emailBody = find(emailBody, JSON.parse(acronyms));
                    $('#post').html(emailBody);
                });
            });
        });
    });
    /** END OF SCRIPT **/
})();