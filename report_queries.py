# Each query should have four parts in its respective list entry:
#   1: Report Name which will also become the output file name
#   2: The query itself which should be held within a docstring
#   3: The name of the fields returned in the query, aliased, blank, or otherwise
#   4: The organizational unit the report belongs to: eg Case Analysts, Managers etc
queries = [
	[
		"Accounts with negative credits",
		"""
            SELECT a.accountNumber, a.disposition, a.credits, t.lastTransactionDate, t.creditsUsed, u.name
            FROM accounts a 
            INNER JOIN transactions t ON t.accountID = a.ID
            INNER JOIN users u ON i.ID = t.userID
            WHERE a.credits < 0
            ORDER BY a.credits 
            
        """,
        ['Account Number', 'Account Status', 'Current Credits', 'Last Transaction Data', 'Credits Used', 'Person Responsible'],
        "Case Analysts"
	],
		
	[
		"Employees who haven't logged in this week",
		"""
            SELECT employee.ID, employee.firstName, employee.lastName, CONVERT(varchar, employee.lastLogin, 0), CONCAT(manager.firstName, " ", manager.lastName) as "Manager Name"
            FROM users employee
            INNER JOIN users manager
            ON manager.id = employee.managerID
            WHERE u.lastLogin < DATEADD(day, -7, GETDATE())
         """,
        ['Employee ID', 'First Name', 'Last Name', 'Last Login', 'Manager Name'],
        "Managers"

	]
]