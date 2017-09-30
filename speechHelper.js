function listCommandList() {
	var ListCommands = new Object();
	ListCommands.noteBooks = "list Notebooks";
	ListCommands.sections = "list Sections";
	ListCommands.pages = "list Pages";
}

function linkAccount() {
	return "Please link your Microsoft account to use this skill";
}

function welcomeMessages() {
	var ListCommands = listCommandList();
	var Welcome = new Object();

	Welcome.message = "Welcome to Graph Bot. . Here are some things you can say. . . " + ListCommands.notebooks + " . . " + ListCommands.sections + " . . or " + ListCommands.pages;
	Welcome.messageCard = 'Here are some things you can say:\n\n ' + ListCommands.noteBooks + ' \n ' + ListCommands.sections + ' \n ' + ListCommands.pages;
	Welcome.messageCardTitle = "Welcome to Graph Bot"

	return Welcome;
}

function shutdownMessage() {
	return "Enjoy your meal.";
}

function helpMessage() {
	var Help = new Object();
	Help.message = "Here are some things you can say. . . Send Mail . . Send a test message . . or Send me mail";
	Help.messageCard = 'Here are some things you can say:\n\n "Send Mail" \n "Send a test message" \n "Send me mail"';
	Help.messageCardTitle = "Graph Bot Help"

	return Help;
}