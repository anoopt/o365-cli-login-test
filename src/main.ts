import * as core from '@actions/core';
import * as exec from '@actions/exec';
import * as io from '@actions/io';

var cliPath: string;

async function main() {
    try{
        let cliInstallCommand: string = "npm install -g @pnp/office365-cli";

        if(process.env.RUNNER_OS == "Windows") {
            await exec.exec(cliInstallCommand);
        } else {
            await exec.exec(`sudo ${cliInstallCommand}`);
        }
        
        cliPath = await io.which("o365", true);
        
        let username = process.env.USERNAME;
        let password = process.env.PASSWORD;
        
        await executeO365CLICommand(`login --authType password --userName ${username} --password ${password}`);
        await executeO365CLICommand("status");
        console.log("Login successful.");
        
    } catch (error) {
        core.error("Login failed. Please check the credentials. For more information refer https://aka.ms/create-secrets-for-GitHub-workflows");
        core.setFailed(error);
    } 
}

async function executeO365CLICommand(command: string) {
    try {
        await exec.exec(`"${cliPath}" ${command}`, [],  {}); 
    }
    catch(error) {
        throw new Error(error);
    }
}

main();