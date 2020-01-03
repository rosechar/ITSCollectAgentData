import json
import boto3

# function to collect info regarding each call center agent and export
# data to a Google Sheet to back up 
def lambda_handler(event, context):
    # initialize connect SDK
    client = boto3.client("connect")

    # get list of agents in our Connect instance
    response = client.list_users(
    InstanceId='********-****-****-****-************',
    MaxResults = 1000
    )

    # extract ID's from response and make list
    listOfIDs = []
    for i in response["UserSummaryList"]:
        listOfIDs.append(i['Id'])

    # initialize Google Sheet
    cellIndex = 2
    client2 = gspread.authorize(creds)
    sheet = client2.open('Connect Users List').get_worksheet(0)
    if creds.access_token_expired:
        client2.login()

    # make API call to collect agent data for each agent
    for i in listOfIDs:
        response = client.describe_user(
        UserId = i,
        InstanceId = 'a351243b-6844-41b0-97d9-934414342d87'
        )
        startCell = 'A' + str(cellIndex)
        endCell = 'L' + str(cellIndex)
        range = startCell + ':' + endCell
        cellIndex = cellIndex + 1
        cell_list = sheet.range(range)
        # extract list of data about agent
        listData = []
        listData.append(response["User"]["Id"])
        listData.append(response["User"]["Arn"])
        listData.append(response["User"]["Username"])
        listData.append(response["User"]["IdentityInfo"]["FirstName"])
        listData.append(response["User"]["IdentityInfo"]["LastName"])
        listData.append(response["User"]["PhoneConfig"]["PhoneType"])
        listData.append(response["User"]["PhoneConfig"]["AutoAccept"])
        listData.append(response["User"]["PhoneConfig"]["AfterContactWorkTimeLimit"])
        listData.append(response["User"]["PhoneConfig"]["DeskPhoneNumber"])
        try:
            listData.append(response["User"]["DirectoryUserId"])
        except KeyError:
            listData.append("")
        for j in response["User"]["SecurityProfileIds"]:
            listData.append(j)
        listData.append(response["User"]["RoutingProfileId"])
        k = 0
        # update row in Google Sheet for individual agent
        for cell in cell_list:
            cell.value = listData[k]
            k += 1
        sheet.update_cells(cell_list)
    return 0
