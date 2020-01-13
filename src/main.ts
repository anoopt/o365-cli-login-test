import * as core from '@actions/core';
import * as exec from '@actions/exec';
import * as io from '@actions/io';
import {chmodSync} from 'fs';

var cliPath: string;

async function main() {
    try{

        await exec.exec("sudo npm install -g @pnp/office365-cli");
        
        cliPath = await io.which("o365", true);
        
        let username = process.env.USERNAME;
        let password = process.env.PASSWORD;
        
        await executeO365CLICommand(`login --authType password --userName ${username} --password ${password}`);
        await executeO365CLICommand("status");
        console.log("Login successful.");

        chmodSync("./scripts/mail.sh", 0o755); 
        await exec.exec('"./scripts/mail.sh"');
        

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