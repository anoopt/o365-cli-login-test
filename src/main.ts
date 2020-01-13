import * as core from '@actions/core';
import * as exec from '@actions/exec';
import * as io from '@actions/io';

let o365CLIPath: string;

async function main() {
    try{
        core.info("Installing Office 365 CLI.");

        let o365CLIInstallCommand: string = "npm install -g @pnp/office365-cli";
        const options: any = {};
        options.silent = true;
        if(process.env.RUNNER_OS == "Windows") {
            await exec.exec(o365CLIInstallCommand, [], options);
        } else {
            await exec.exec(`sudo ${o365CLIInstallCommand}`, [], options);
        }
        o365CLIPath = await io.which("o365", true);
        
        core.info("Completed installing Office 365 CLI.");

        core.info("Logging in to the tenant.");

        const username: string = core.getInput("USERNAME");
        const password: string = core.getInput("PASSWORD");
        await executeO365CLICommand(`login --authType password --userName ${username} --password ${password}`);
        await executeO365CLICommand("status");

        core.info("Login successful.");
        
    } catch (error) {
        core.error("Login to the tenant failed. Please check the credentials. For more information refer https://aka.ms/create-secrets-for-GitHub-workflows");
        core.setFailed(error);
    } 
}

async function executeO365CLICommand(command: string) {
    try {
        await exec.exec(`"${o365CLIPath}" ${command}`, [],  {}); 
    }
    catch(error) {
        throw new Error(error);
    }
}

main();