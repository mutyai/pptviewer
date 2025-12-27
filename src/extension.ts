import * as vscode from 'vscode';
import * as path from 'path';
import { PPTEditorProvider } from './ppt-editor-provider';

export function activate(context: vscode.ExtensionContext) {
	console.log('Muty PPT Viewer extension is now active!');

	// Register custom editor provider for PPT files
	const provider = new PPTEditorProvider(context.extensionUri);
	context.subscriptions.push(
		vscode.window.registerCustomEditorProvider('muty-pptviewer.preview', provider, {
			supportsMultipleEditorsPerDocument: false,
			webviewOptions: {
				retainContextWhenHidden: true
			}
		})
	);

	// Register the command to open PPT files
	const openPPTCommand = vscode.commands.registerCommand('muty-pptviewer.openPPT', async (uri: vscode.Uri) => {
		if (!uri) {
			vscode.window.showErrorMessage('Please select a PowerPoint file to preview.');
			return;
		}

		// Check if file is a PowerPoint file
		const ext = path.extname(uri.fsPath).toLowerCase();
		if (ext !== '.pptx' && ext !== '.ppt') {
			vscode.window.showErrorMessage('Please select a valid PowerPoint file (.pptx or .ppt).');
			return;
		}

		// Open the file with our custom editor
		await vscode.commands.executeCommand('vscode.openWith', uri, 'muty-pptviewer.preview');
	});

	context.subscriptions.push(openPPTCommand);
}

export function deactivate() { }
