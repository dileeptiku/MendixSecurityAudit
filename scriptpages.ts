
import { MendixSdkClient, OnlineWorkingCopy, Project, Revision, Branch, loadAsPromise } from "mendixplatformsdk";
import { ModelSdkClient, IModel, projects, domainmodels, microflows, pages, navigation, texts, security, IStructure, menus } from "mendixmodelsdk";


import when = require('when');


const username = "dileep.tiku@mendix.com";
const apikey = "b06e5563-9ea2-4f99-b36f-60b6545942a6";
const projectId = "999f5f7d-6f90-49db-a58f-5d2e6210c656";
const projectName = "USA-eXpTransactions";
const revNo = -1; // -1 for latest
const branchName = null // null for mainline
const wc = null;
const client = new MendixSdkClient(username, apikey);
var officegen = require('officegen');
var xlsx = officegen('xlsx');
var fs = require('fs');
var pObj;

const sheet = xlsx.makeNewSheet ();
sheet.name = 'Entities';

sheet.data[0]=[];
sheet.data[0][0] = `User Role`;
sheet.data[0][1] = `Module`;
sheet.data[0][2] = `Module Role`;
sheet.data[0][3] = `Entity`;
sheet.data[0][4] = `Xpath`;
sheet.data[0][5] = `Create/Delete`;
sheet.data[0][6] = `Member Rules`;


const sheetPages = xlsx.makeNewSheet ();
sheetPages.name = 'Pages';

sheetPages.data[0]=[];
sheetPages.data[0][0] = `User Role`;
sheetPages.data[0][1] = `Module`;
sheetPages.data[0][2] = `Module Role`;
sheetPages.data[0][3] = `Page Name`;
sheetPages.data[0][4] = `Allowed`;

  
/*
 * PROJECT TO ANALYZE
 */
const project = new Project(client, projectId, projectName);
main();


async function main(){

    const workingCopy = await loadWorkingCopy(project, new Revision(revNo, new Branch(project, branchName)));

    const projectSecurity = await loadProjectSecurity(workingCopy);

    const userRoles = await getAllUserRoles(projectSecurity);
    
    const securityDocument = await createUserSecurityDocument(userRoles);

    var out = fs.createWriteStream('MendixSecurityDocument.xlsx');
    xlsx.generate(out);
    out.on('close', function () {
        console.log('Finished to creating Document');
    });


}

function loadWorkingCopy(project:Project, revision:Revision):when.Promise<OnlineWorkingCopy>{
    return client.platform().createOnlineWorkingCopy(project, revision);
}

/**
* This function picks the first navigation document in the project.
*/
function createUserSecurityDocument(userRoles: security.UserRole[]): when.Promise<security.UserRole[]> {
    return when.all<security.UserRole[]>(userRoles.map(addText));
}

function addText(userRole: security.UserRole): when.Promise<void> {
    return processUsersSecurity(userRole);
}

function processUsersSecurity(userRole: security.UserRole): when.Promise<void> {
    console.log(`Processing User Role: ${userRole.name}`)
    return processAllModules(userRole.model.allModules(), userRole);
    //return processAllModules(userRole.model.allModules().filter(module => module.name === "Transaction"), userRole);
    
}
function processAllModules(modules: projects.IModule[], userRole: security.UserRole): when.Promise<void> {
    return when.all<void>(modules.map(module => processModule(module, userRole)))
}

function processModule(module: projects.IModule, userRole: security.UserRole): when.Promise<void> {
    console.log(`Processing module: ${module.name}`);
    var securities = getAllModuleSecurities(module);
    return when.all<void>(securities.map(security => loadAllModuleSecurities(securities,userRole)));
    
}

function loadAllModuleSecurities(moduleSecurities: security.IModuleSecurity[], userRole: security.UserRole): when.Promise<void> {
    return when.all<void>(moduleSecurities.map(mSecurity => processLoadedModSec(mSecurity,userRole)));
}

function getAllModuleSecurities(module: projects.IModule): security.IModuleSecurity[] {
    console.log(`Processing getAllModuleSecurities: ${module.name}`);
    return module.model.allModuleSecurities().filter(modSecurity => {
        if (modSecurity != null) {
			console.log(`Mod Security is not null: ${modSecurity.containerAsModule.name}`);
            return modSecurity.containerAsModule.name === module.name;
        } else {
            return false;
        };

    });
}

function loadModSec(modSec: security.IModuleSecurity): when.Promise<security.ModuleSecurity> {
    console.log(`Processing loadModSec`);
    return loadAsPromise(modSec);
}

function processLoadedModSec(modSec: security.IModuleSecurity, userRole: security.UserRole, ):when.Promise<void>{
    return when.all<void>(modSec.moduleRoles.map(modRole => processModRole(modRole,userRole)));
}

function processModRole(modRole:security.IModuleRole, userRole:security.UserRole):when.Promise<void>{
    if(addIfModuleRoleInUserRole(modRole, userRole)){
        return detailEntitySecurity(modRole,userRole);
    }
    return when.resolve();
}

function detailEntitySecurity(modRole:security.IModuleRole,userRole:security.UserRole):when.Promise<void>{
    return when.all<void>(modRole.containerAsModuleSecurity.containerAsModule.domainModel.entities.map(entity =>
        processAllEntitySecurityRules(entity,modRole,userRole))).then(()=> processAllPages(modRole,userRole));
}


function processAllPages(modRole:security.IModuleRole,userRole:security.UserRole):when.Promise<void>{
    return when.all<void>(modRole.model.allPages().map(page => processPage(modRole,userRole,page)));
}

function processPage(modRole:security.IModuleRole, userRole:security.UserRole, page:pages.IPage):when.Promise<void>{
        return loadAsPromise(page).then(loadedPage =>addPage(modRole,userRole,loadedPage));       
}

function addPage(modRole:security.IModuleRole, userRole:security.UserRole, loadedPage:pages.Page){
	if(loadedPage.allowedRoles.filter(allowedRole => allowedRole?.name == modRole.name).length > 0){
		sheetPages.data.push([`${userRole.name}`,`${modRole.containerAsModuleSecurity.containerAsModule.name}`,`${modRole.name}`,`${loadedPage.name}`,`True`]);
		//console.log(`${userRole.name}`,`${modRole.containerAsModuleSecurity.containerAsModule.name}`,`${modRole.name}`,`${loadedPage.name}`,`True`);
		console.log(`Add page: ${modRole.name}`,`${userRole.name}`,`${modRole.containerAsModuleSecurity.containerAsModule.name}`);
	}else{
		sheetPages.data.push([`${userRole.name}`,`${modRole.containerAsModuleSecurity.containerAsModule.name}`,`${modRole.name}`,`${loadedPage.name}`,`False`]);
		//console.log(`${userRole.name}`,`${modRole.containerAsModuleSecurity.containerAsModule.name}`,`${modRole.name}`,`${loadedPage.name}`,`False`);
	}
}



 function processAllEntitySecurityRules(entity:domainmodels.IEntity,moduleRole:security.IModuleRole,userRole:security.UserRole):when.Promise<void>{
    return loadAsPromise(entity).then(loadedEntity => 
        checkIfModuleRoleIsUsedForEntityRole(loadedEntity,loadedEntity.accessRules, moduleRole,userRole));
}

function checkIfModuleRoleIsUsedForEntityRole(entity:domainmodels.Entity,accessRules:domainmodels.AccessRule[], modRole:security.IModuleRole,userRole:security.UserRole):when.Promise<void>{
    return when.all<void>(
        accessRules.map(rule =>{
            var memberRules = ``;
            if(rule.moduleRoles.filter(entityModRule =>{
                return entityModRule.name === modRole.name;
            }).length > 0){
                    rule.memberAccesses.map(memRule =>{
                        if(memRule != null){
                            if(memRule.accessRights!= null && memRule.attribute != null){
                                memberRules += `${memRule.attribute.name}: ${memRule.accessRights.name}\n`;
                            }
                        }
                        
                    }
                );
                var createDelete;
                if(rule.allowCreate && rule.allowDelete){
                    createDelete = `Create/Delete`
                 }else if(rule.allowCreate){
                    createDelete = `Create`
                 }else if(rule.allowDelete){
                    createDelete = `Delete`
                 }else{
                    createDelete = `None`
                 }
                sheet.data.push([`${userRole.name}`,`${entity.containerAsDomainModel.containerAsModule.name}`,`${modRole.name}`,`${entity.name}`,`${rule.xPathConstraint}`,`${createDelete}`,`${memberRules}`]);
                //console.log(`${userRole.name},${entity.containerAsDomainModel.containerAsModule.name},${modRole.name},${entity.name},${rule.xPathConstraint},${createDelete},${memberRules}`);
            }
        })
    );

} 

function addIfModuleRoleInUserRole(modRole: security.IModuleRole, userRole: security.UserRole): boolean{
        //console.log(`Processing module role: ${modRole.name}`);
        if (userRole.moduleRoles.filter(modRoleFilter => {
            if (modRoleFilter != null) {
                return modRoleFilter.name === modRole.name;
            } else {
                return false;
            }
        }).length > 0) {
            return true;
        }else{
            return false;
        }
        
}

function getAllModules(workingCopy: OnlineWorkingCopy): projects.IModule[] {
    return workingCopy.model().allModules().filter(module => module.name==="Finance");
}

function processDomainModel(module: projects.IModule, role: security.UserRole): when.Promise<void> {
    return when.all<void>(module.domainModel.entities.map((entity) => checkEntity(entity)));
}

function checkEntity(entity: domainmodels.IEntity) {
    return loadAsPromise(entity).then(ent => {
        ent.accessRules
    });
}

/**
* This function loads the project security.
*/
function loadProjectSecurity(workingCopy: OnlineWorkingCopy): when.Promise<security.ProjectSecurity> {
    var security = workingCopy.model().allProjectSecurities()[0];
    return when.promise<security.ProjectSecurity>((resolve, reject) => {
        if (security) {
            security.load(secure => {
                if (secure) {
                    console.log(`Loaded security`);
                    resolve(secure);
                } else {
                    console.log(`Failed to load security`);
                    reject(`Failed to load security`);
                }
            });
        } else {
            reject(`'security' is undefined`);
        }
    });
}

function getAllUserRoles(projectSecurity: security.ProjectSecurity): security.UserRole[] {
    return projectSecurity.userRoles;
}
