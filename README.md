# TeamsTimecardHelper
Summarize and edit Teams Timecards

This is the settings used to initialize this repo

    dotnet new blazorwasm -o TeamsTimecardHelperClient -f net7.0 -au Individual --client-id 80dd4c35-6e75-4ec5-96ee-d97f88d077d8 --calls-graph 

As well as this to get the AzureAD/MSAL part after commiting

    dotnet new blazorwasm -o TeamsTimecardHelperClient -f net7.0 -au SingleOrg --client-id 80dd4c35-6e75-4ec5-96ee-d97f88d077d8 --calls-graph --force
