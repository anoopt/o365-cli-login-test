import {getInput, info, error, setFailed, setOutput} from '@actions/core';
import { exec } from '@actions/exec';
import { which } from '@actions/io';

let o365CLIPath: string;

async function main() {
    try{
        info("Installing Office 365 CLI.");

        let o365CLIInstallCommand: string = "npm install -g @pnp/office365-cli";
        const options: any = {};
        options.silent = true;
        if(process.env.RUNNER_OS == "Windows") {
            await exec(o365CLIInstallCommand, [], options);
        } else {
            await exec(`sudo ${o365CLIInstallCommand}`, [], options);
        }
        o365CLIPath = await which("o365", true);
        
        info("Completed installing Office 365 CLI.");

        info("Logging in to the tenant...");

        const username: string = getInput("USERNAME");
        const password: string = getInput("PASSWORD");
        await executeO365CLICommand(`login --authType password --userName ${username} --password ${password}`);
        await executeO365CLICommand("status");

        info("Login successful.");
        
    } catch (error) {
        error("Login to the tenant failed. Please check the credentials. For more information refer https://aka.ms/create-secrets-for-GitHub-workflows");
        setFailed(error);
    } 
}

async function executeO365CLICommand(command: string) {
    try {
        await exec(`"${o365CLIPath}" ${command}`, [],  {}); 
    }
    catch(error) {
        throw new Error(error);
    }
}

main();