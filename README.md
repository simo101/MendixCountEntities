# Mendix Count Entities Script

This script creates a word document with each module and the Entities within each module.

## Setup
To set up and use the sdk you need to have node js installed on your machine. You will need to have also installed typescript and tsd.
The following command will install typescript and tsd globally for you:

`npm install -g typescript`

Open up the folder using node.js.
To install the count entities script you should type the command:

`npm install`

This is will install the script and all the relevant dependencies.

To connect it to your project you need to change the following constants in the `script.ts`


`var appID = "{{appID}}";`

`var projectName = "{{ProjectName}}";`

API keys can be found in the mendix home portal.

## Use the Script
Once the visualiser is installed type:
`tsc`
to compile the script code.

Then after compiled type:
`node script.js`