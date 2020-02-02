# OutlookSearcher

OutlookSearcher uses COM to pull data from Outlook's MAPI namespace. Outlook must be running for this tool to work.

Compatible with cobalt strike

### Usage:
    Arguments:
        Required:
            searchterms     Specify a comma deliminated list of searchterms. This will search through every email in every Outlook folder.

### Examples:
            OutlookSearcher.exe searchterms=password
            OutlookSearcher.exe searchterms=password,asfd
            
            Output:
                From: person@example.com
                To: asdf123@hotmail.com
                Subject: here's that thing!
                Hey so don't forget the password is 123

### Dependencies:
        .Net Framework 4.0+