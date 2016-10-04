/// <reference path='./typings/tsd.d.ts' />

'use strict';

import { MendixSdkClient, OnlineWorkingCopy, Project, Revision, Branch, loadAsPromise } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus } from "mendixmodelsdk";


import when = require('when');


const username = "{{UserName}}";
const apikey = "{{APIKey}}";
const projectId = "{{ProjectID}}";
const projectName = "{{ProjectName}}";
const revNo = -1; // -1 for latest
const branchName = null // null for mainline
const wc = null;
const client = new MendixSdkClient(username, apikey);
var officegen = require('officegen');
var docx = officegen('docx');
var fs = require('fs');
var pObj;
/*
 * PROJECT TO ANALYZE
 */
const project = new Project(client, projectId, projectName);

client.platform().createOnlineWorkingCopy(project, new Revision(revNo, new Branch(project, branchName)))
    .then(workingCopy => {
        pObj = docx.createP();
        workingCopy.model().allDomainModels().forEach(domainModel=>{
                pObj.addText(domainModel.moduleName, { bold: true, underline: true, font_size: 20 });
                pObj.addLineBreak();


                pObj.addText(`Total Entities: ${domainModel.entities.length}`, { bold: true, underline: false, font_size: 18 });
                pObj.addLineBreak();              

                pObj.addText(`Entity Names:`, { bold: true, underline: false, font_size: 16 });
                pObj.addLineBreak();              
                
                domainModel.entities.forEach(entity =>{
                    pObj.addText(entity.name, { bold: false, underline: false, font_size: 15 });
                    pObj.addLineBreak();
                });
                pObj.addLineBreak(); 
                                
                                
                var totalPages = workingCopy.model().allPages().filter(page =>{
                    return page.moduleName === domainModel.moduleName;
                });

                pObj.addText(`Total Pages: ${totalPages.length}`, { bold: true, underline: false, font_size: 18 });
                pObj.addLineBreak();          

                pObj.addText(`Pages:`, { bold: true, underline: false, font_size: 16 });
                pObj.addLineBreak(); 
                totalPages.forEach(pg =>{
                    pObj.addText(pg.name, { bold: false, underline: false, font_size: 15 });
                    pObj.addLineBreak();
                });

                pObj.addLineBreak();
                var microflows = workingCopy.model().allMicroflows().filter(microflow =>{
                    return microflow.moduleName === domainModel.moduleName;
                });        

                pObj.addText(`Total Microflows: ${microflows.length}`, { bold: true, underline: false, font_size: 18 });
                pObj.addLineBreak();       

                pObj.addText(`Microflows:`, { bold: true, underline: false, font_size: 16 });
                pObj.addLineBreak(); 
                microflows.forEach(mf =>{
                    pObj.addText(mf.name, { bold: false, underline: false, font_size: 15 });
                    pObj.addLineBreak();
                });


                return;
        });
        return;
        })
    .done(
    () => {
        var out = fs.createWriteStream('MendixCountDocument.docx');
        docx.generate(out);
        out.on('close', function () {
            console.log('Finished to creating Document');
        });
    },
    error => {
        console.log("Something went wrong:");
        console.dir(error);
    }
    );