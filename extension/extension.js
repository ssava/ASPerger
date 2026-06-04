const vscode = require('vscode');

function activate(context) {
    context.subscriptions.push(
        vscode.debug.registerDebugAdapterDescriptorFactory('asperger', {
            createDebugAdapterDescriptor(session) {
                const command = context.asAbsolutePath('./bin/asperger-debug');
                return new vscode.DebugAdapterExecutable(command);
            }
        })
    );
}

function deactivate() {}

module.exports = { activate, deactivate };
