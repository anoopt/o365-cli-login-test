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
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
const core = __importStar(require("@actions/core"));
const exec = __importStar(require("@actions/exec"));
const io = __importStar(require("@actions/io"));
let o365CLIPath;
function main() {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            core.info("Installing Office 365 CLI.");
            let o365CLIInstallCommand = "npm install -g @pnp/office365-cli";
            const options = {};
            options.silent = true;
            if (process.env.RUNNER_OS == "Windows") {
                yield exec.exec(o365CLIInstallCommand, [], options);
            }
            else {
                yield exec.exec(`sudo ${o365CLIInstallCommand}`, [], options);
            }
            o365CLIPath = yield io.which("o365", true);
            core.info("Completed installing Office 365 CLI.");
            core.info("Logging in to the tenant...");
            const username = core.getInput("USERNAME");
            const password = core.getInput("PASSWORD");
            yield executeO365CLICommand(`login --authType password --userName ${username} --password ${password}`);
            yield executeO365CLICommand("status");
            core.info("Login successful.");
        }
        catch (error) {
            core.error("Login to the tenant failed. Please check the credentials. For more information refer https://aka.ms/create-secrets-for-GitHub-workflows");
            core.setFailed(error);
        }
    });
}
function executeO365CLICommand(command) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            yield exec.exec(`"${o365CLIPath}" ${command}`, [], {});
        }
        catch (error) {
            throw new Error(error);
        }
    });
}
main();
