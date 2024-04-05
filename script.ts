import { MendixPlatformClient} from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus } from "mendixmodelsdk";

const appID = "fcc8a372-b054-4cb8-876c-ba4beeb7b41b";
const projectName = "Mendix Experience";
const revNo = -1; // -1 for latest
const branchName = null // null for mainline
const wc = null;
const client = new MendixPlatformClient();
var officegen = require('officegen');
var docx = officegen('docx');
var fs = require('fs');
var pObj = docx.createP();
var totalNumberPages = 0;
var totalNumberMicroflows=0;
var totalNumberEntities = 0;


/*
 * PROJECT TO ANALYZE
 */
async function main() {
    const app = client.getApp(appID);
    var useBranch:string ="";
    if(branchName === null){
        var repositoryInfo = await app.getRepository().getInfo();
        if (repositoryInfo.type === `svn`)
            useBranch = `trunk`;
        else
            useBranch = `main`;
    }else{
        useBranch = branchName;
    }
    const workingCopy = await app.createTemporaryWorkingCopy(useBranch);
        pObj.addText(projectName, { bold: true, underline: true, font_size: 20 });
        pObj.addLineBreak();
        pObj.addLineBreak();
        const model = await workingCopy.openModel();
        model.allDomainModels().forEach(domainModel => {
            pObj.addText(getModuleName(domainModel), { bold: true, underline: true, font_size: 18 });
            pObj.addLineBreak();

            totalNumberEntities+= domainModel.entities.length;
            pObj.addText(`Total Entities: ${domainModel.entities.length}`, { bold: false, underline: false, font_size: 15 });
            pObj.addLineBreak();

            var totalPages = model.allPages().filter(page => {
                return getModuleName(page) === getModuleName(domainModel);
            });
            totalNumberPages+= totalPages.length;
            pObj.addText(`Total Pages: ${totalPages.length}`, { bold: false, underline: false, font_size: 15 });

            pObj.addLineBreak();
            var microflows = model.allMicroflows().filter(microflow => {
                return getModuleName(microflow) === getModuleName(domainModel);
            });
            totalNumberMicroflows+= microflows.length;
            pObj.addText(`Total Microflows: ${microflows.length}`, { bold: false, underline: false, font_size: 15 });
            pObj.addLineBreak();
            pObj.addLineBreak();
            
            return;
        });
        pObj.addText(`Total Stats`, { bold: true, underline: true, font_size: 18 });
        pObj.addLineBreak();
        pObj.addText(`Total Application Objects: ${totalNumberPages+totalNumberEntities+totalNumberMicroflows}`, { bold: false, underline: false, font_size: 15 });
        pObj.addLineBreak();
        pObj.addText(`Total Pages: ${totalNumberPages}`, { bold: false, underline: false, font_size: 15 });
        pObj.addLineBreak();
        pObj.addText(`Total Microflows: ${totalNumberMicroflows}`, { bold: false, underline: false, font_size: 15 });
        pObj.addLineBreak();
        pObj.addText(`Total Entities: ${totalNumberEntities}`, { bold: false, underline: false, font_size: 15 });
        
        var out = fs.createWriteStream(`${projectName} Application Counts.docx`);
        docx.generate(out);
        out.on('close', function () {
            console.log('Finished to creating Document');
        });
    }
    export function getModuleName(element: IStructure): string {
        let current = element.unit;
        while (current) {
            if (current instanceof projects.Module) {
                return current.name;
            }
            current = current.container;
        }
        return "";
    }
main().catch(console.error);