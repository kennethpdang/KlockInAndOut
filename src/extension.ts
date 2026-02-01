import * as vscode from 'vscode';
import * as path from 'path';

export function activate(context: vscode.ExtensionContext) {
    const disposable = vscode.commands.registerCommand(
        'kendang.clockInOut',
        async () => {
            // Get the workspace folder
            const workspaceFolders = vscode.workspace.workspaceFolders;
            if (!workspaceFolders || workspaceFolders.length === 0) {
                vscode.window.showErrorMessage('No workspace folder open. Please open a folder first.');
                return;
            }

            const workspacePath = workspaceFolders[0].uri.fsPath;
            const scriptPath = path.join(context.extensionPath, 'timesheet.py');

            // Create or reuse terminal
            let terminal = vscode.window.terminals.find(t => t.name === 'Timesheet');
            if (!terminal) {
                terminal = vscode.window.createTerminal('Timesheet');
            }

            terminal.show(true);
            terminal.sendText(`python "${scriptPath}" "${workspacePath}"`);
        }
    );

    context.subscriptions.push(disposable);
}

export function deactivate() {}