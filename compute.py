import pandas as pd

# Provide the file paths for the DEV and SIT sheets
filePath = 'UserGroups.xlsx'
outputFilePath = 'output.xlsx'

# Flags for export function
comparedUserNames = 0
foundMissingGroups = 0
foundMissingGroupsPerUser = 0
foundMissingPermissions = 0

def findMissingGroups(filePath):
    # Read the DEV and SIT sheets, considering only columns 1, 3, and 4
    devData = pd.read_excel(filePath, sheet_name=0, usecols=[1, 3, 4])
    sitData = pd.read_excel(filePath, sheet_name=1, usecols=[1, 3, 4])

    # Convert NaN values to empty strings in the 'Company Name' column
    devData['Company Name'] = devData['Company Name'].fillna('')
    sitData['Company Name'] = sitData['Company Name'].fillna('')

    # Convert the data to sets for easier comparison
    devGroups = set(zip(devData['User Name'], devData['User Group Code'], devData['Company Name']))
    sitGroups = set(zip(sitData['User Name'], sitData['User Group Code'], sitData['Company Name']))

    # Find missing groups in DEV and SIT sheets
    missingDevGroups = sitGroups - devGroups
    missingSitGroups = devGroups - sitGroups

    # Prepare the output data
    outputData = []
    for user, group, company in missingDevGroups:
        outputData.append({'User Name': user, 'Environment': 'DEV', 'Group Code': group, 'Company Name': company})
    for user, group, company in missingSitGroups:
        outputData.append({'User Name': user, 'Environment': 'SIT', 'Group Code': group, 'Company Name': company})

    # Create DataFrame from the output data
    df = pd.DataFrame(outputData)

    # foundMissingGroups = 1 - foundMissingGroups

    return df

def compareUserNames(filePath):
    devData = pd.read_excel(filePath, sheet_name=0, usecols=[4])
    sitData = pd.read_excel(filePath, sheet_name=1, usecols=[4])

    # Find unique user names in both DEV and SIT sheets
    devUserNames = set(devData['User Name'])
    sitUserNames = set(sitData['User Name'])

    # Determine the environment for each user
    users = []
    for user in devUserNames.union(sitUserNames):
        if user in devUserNames and user in sitUserNames:
            environment = 'DEV AND SIT'
        elif user in devUserNames:
            environment = 'DEV'
        else:
            environment = 'SIT'
        users.append({'User Name': user, 'Environment': environment})

    # Create DataFrame from the output data
    df = pd.DataFrame(users)

    # comparedUserNames = 1 - comparedUserNames

    return df

def missingGroupsUsersExistInBoth(comparedUsersDf, missingGroupsDf):
    bothEnvUsers = comparedUsersDf.loc[comparedUsersDf['Environment'] == 'DEV AND SIT', 'User Name'].tolist()

    # missingGroupsBothEnv = missingGroupsDf[
    #     (missingGroupsDf['User Name'].isin(bothEnvUsers)) &
    #     (missingGroupsDf['Environment'] == 'SIT')
    # ]
    missingGroupsBothEnv = missingGroupsDf[
        missingGroupsDf['User Name'].isin(bothEnvUsers)
    ]

    df = pd.DataFrame(missingGroupsBothEnv)

    return df



def exportToExcel(comparedUsersDf, missingGroupsDf, missingGroupsPerUserDf, outputFilePath):
    with pd.ExcelWriter(outputFilePath, engine='openpyxl', mode='w') as writer:
        comparedUsersDf.to_excel(writer, sheet_name='User Environment Mapping', index=False)
        missingGroupsDf.to_excel(writer, sheet_name='Users Missing Groups', index=False)
        missingGroupsPerUserDf.to_excel(writer, sheet_name='Missing Groups Per User', index=False)



# Compare user names and create DataFrame
comparedUsersDf = compareUserNames(filePath)

# Find missing groups and create DataFrame
missingGroupsDf = findMissingGroups(filePath)

missingGroupsPerUserDf = missingGroupsUsersExistInBoth(comparedUsersDf, missingGroupsDf)

exportToExcel(comparedUsersDf, missingGroupsDf, missingGroupsPerUserDf, outputFilePath)