function listCommandList() {
	var ListCommands = new Object();
	ListCommands.noteBooks = "list Notebooks";
	ListCommands.sections = "list Sections";
	ListCommands.pages = "list Pages";

	return ListCommands;
}

function linkAccount() {
	return "Please link your Microsoft account to use this skill";
}

function welcomeMessages() {
	var ListCommands = listCommandList();
	var Welcome = new Object();

	Welcome.message = "Welcome to Recipe Notes. . Here are some things you can say. . . " + ListCommands.noteBooks + " . . " + ListCommands.sections + " . . or " + ListCommands.pages;
	Welcome.messageCard = 'Here are some things you can say:\n\n ' + ListCommands.noteBooks + ' \n ' + ListCommands.sections + ' \n ' + ListCommands.pages;
	Welcome.messageCardTitle = "Welcome to Recipe Notes"

	return Welcome;
}

function shutdownMessage() {
	return "Enjoy your meal.";
}

function helpMessage() {
	var Help = new Object();
	Help.message = "Here are some things you can say. . . " + ListCommands.noteBooks + " . . " + ListCommands.sections + " . . or " + ListCommands.pages";
	Help.messageCard = 'Here are some things you can say:\n\n ' + ListCommands.noteBooks + ' \n ' + ListCommands.sections + ' \n ' + ListCommands.pages;
	Help.messageCardTitle = "Recipe Notes Help"

	return Help;
}

function listItemTypes(itemType) {
	return "Your " + itemType + " are " + sayArray(listOfSections,  'and')
                    + '. please choose either ' +  sayArray(listOfSections,  'or');
}

function sayArray(myData, andor) {
    // the first argument is an array [] of items
    // the second argument is the list penultimate word; and/or/nor etc.

    var listString = '';

    if (myData.length == 1) {
        listString = myData[0];
    } else {
        if (myData.length == 2) {
            listString = myData[0] + ' ' + andor + ' ' + myData[1];
        } else {

            for (var i = 0; i < myData.length; i++) {
                if (i < myData.length - 2) {
                    listString = listString + myData[i] + ', ';
                    if (i = myData.length - 2) {
                        listString = listString + myData[i] + ', ' + andor + ' ';
                    }

                } else {
                    listString = listString + myData[i];
                }
            }
        }
    }

    return(listString);
}
