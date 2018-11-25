## SPFxDemo Azure Function Azure AD
This is the sample I use in my blogpost [Part 4 - Azure Functions V2 + VS Code + PnPJs === true](https://simonagren.github.io/part4-azurefunction/)

We call an Azure AD secured Azure Function, and receive the lists from a site.

### Minimal Path to Awesome

```bash
git clone the repo
npm i 
```
- Change the <ApplicationId>
- Change to your Site name
- Change to you FunctionApp Url
- Package and bundle the app
```bash
gulp bundle --ship && gulp package-solution --ship
```
- Upload to ypu app catalog and deploy
- Add HelloWorld web part to one of your pages.
