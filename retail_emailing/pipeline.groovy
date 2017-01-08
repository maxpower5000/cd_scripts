def userInput
def workspace

stage('User input') {
	timeout(time: 1, unit: 'DAYS') {
		userInput = input(
			id: 'u_in', message: 'Collect user input', parameters: [
				string(defaultValue: '', description: 'User name, ex. ivanovii', name: 'Username'),
				password(defaultValue: '', description: 'User password', name: 'Password'),
				choice(choices: 'RETAIL\nNA\nDIAUP\nFACT\n', description: 'Project key in Jira', name: 'Project key'),
				string(defaultValue: '', description: 'Patch version', name: 'Version')
			]
		)
	}
}

stage('Initiating') {
	node {
		workspace = 'C:\\Users\\NEX\\Documents\\jenkins_pipeline\\retail_emailing'
		bat "PowerShell.exe -ExecutionPolicy ByPass -Command \"& \'${workspace}\\initiating.ps1\' -WorkSpace \'${workspace}\'\""
	}
}

stage('Get Jira issues') {
	node {
		bat (
			"@PowerShell.exe -ExecutionPolicy ByPass -Command \"& \'${workspace}\\get_jira_issues.ps1\' " +
			"-ProjKey \'${userInput['Project key']}\' -ProjVer \'${userInput['Version']}\' -UserName \'${userInput['Username']}\' " +
			"-Password \'${userInput['Password']}\' -WorkSpace \'${workspace}\'\""
		)
	}
}