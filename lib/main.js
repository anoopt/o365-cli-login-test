"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const core_1 = require("@actions/core");
const exec_1 = require("@actions/exec");
const io_1 = require("@actions/io");
let o365CLIPath;
function main() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            core_1.info("Installing Office 365 CLI.");
            let o365CLIInstallCommand = "npm install -g @pnp/office365-cli";
            const options = {};
            options.silent = true;
            if (process.env.RUNNER_OS == "Windows") {
                yield exec_1.exec(o365CLIInstallCommand, [], options);
            }
            else {
                yield exec_1.exec(`sudo ${o365CLIInstallCommand}`, [], options);
            }
            o365CLIPath = yield io_1.which("o365", true);
            core_1.info("Completed installing Office 365 CLI.");
            core_1.info("Logging in to the tenant...");
            const username = core_1.getInput("USERNAME");
            const password = core_1.getInput("PASSWORD");
            yield executeO365CLICommand(`login --authType password --userName ${username} --password ${password}`);
            yield executeO365CLICommand("status");
            core_1.info("Login successful.");
        }
        catch (error) {
            error("Login to the tenant failed. Please check the credentials. For more information refer https://aka.ms/create-secrets-for-GitHub-workflows");
            core_1.setFailed(error);
        }
    });
}
function executeO365CLICommand(command) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            yield exec_1.exec(`"${o365CLIPath}" ${command}`, [], {});
        }
        catch (error) {
            throw new Error(error);
        }
    });
}
main();
