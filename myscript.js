(function () {
	'use strict';

	// The Office initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			displayItemDetails();
		});
	};

	function getTaskSuggestions() {
		var entities = Office.context.mailbox.item.getEntities();
		var tasksArray = entities.taskSuggestions;
		var htmlText = "";
		// Iterates through each instance of a task suggestion.
		for (var i = 0; i < tasksArray.length; i++)
		{
        // Gets the string that was identified as a task suggestion.
        htmlText += "TaskString : <span>" + 
        tasksArray[i].taskString + "</span><br/>";

        // Gets an array of assignees for that instance of a task 
        // suggestion. Each assignee is represented by an 
        // EmailUser object.
        var assigneesArray = tasksArray[i].assignees;
        for (var j = 0; j < assigneesArray.length; j++)
        {
        	htmlText += "Assignee : ( ";
            // Gets the displayName property of the assignee.
            htmlText += "displayName = <span>" + assigneesArray[j].displayName + 
            "</span> , ";

            // Gets the emailAddress property of each assignee.
            // This is the SMTP address of the assignee.
            htmlText += "emailAddress = <span>" + assigneesArray[j].emailAddress + 
            "</span>";

            htmlText += " )<br/>";
		}

		htmlText += "<hr/>";
		}

		document.getElementById("entities_box").innerHTML = htmlText;
	}




	// Displays the "Subject" and "From" fields, based on the current mail item
	function displayItemDetails() {
		var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
		$('#subject').text(item.subject);

		var from;
		if (item.itemType === Office.MailboxEnums.ItemType.Message) {
			from = Office.cast.item.toMessageRead(item).from;
		} else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
			from = Office.cast.item.toAppointmentRead(item).organizer;
		}

		if (from) {
			$('#from').text(from.displayName);
			$('#from').click(function () {
				app.showNotification(from.displayName, from.emailAddress);
			});
		}
	}
})();