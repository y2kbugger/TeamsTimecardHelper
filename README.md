# TeamsTimecardHelper
Summarize and edit Teams Timecards

This is the settings used to initialize this repo

    dotnet new blazorwasm -o TeamsTimecardHelperClient -f net7.0 -au Individual --client-id 80dd4c35-6e75-4ec5-96ee-d97f88d077d8 --calls-graph

As well as this to get the AzureAD/MSAL part after commiting

    dotnet new blazorwasm -o TeamsTimecardHelperClient -f net7.0 -au SingleOrg --client-id 80dd4c35-6e75-4ec5-96ee-d97f88d077d8 --calls-graph --force

# Development

    cd TeamsTimecardHelperClient
    dotnet watch

When developing on Radzen, dot a `dotnet build` in the Radzen folder at least once to make the css available.

To deploy from CLI, goto the project root and run:

    npx @azure/static-web-apps-cli build
    npx @azure/static-web-apps-cli deploy --env Test
