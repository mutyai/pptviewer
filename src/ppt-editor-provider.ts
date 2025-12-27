import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import { LibreOfficeConverter } from './libreoffice-converter';

export class PPTEditorProvider implements vscode.CustomReadonlyEditorProvider {
	public static readonly viewType = 'muty-pptviewer.preview';

	constructor(private readonly _extensionUri: vscode.Uri) { }

	public async openCustomDocument(
		uri: vscode.Uri,
		openContext: vscode.CustomDocumentOpenContext,
		_token: vscode.CancellationToken,
	): Promise<vscode.CustomDocument> {
		return {
			uri,
			dispose: () => { }
		};
	}

	public async resolveCustomEditor(
		document: vscode.CustomDocument,
		webviewPanel: vscode.WebviewPanel,
		_token: vscode.CancellationToken,
	): Promise<void> {
		// Set up webview
		webviewPanel.webview.options = {
			enableScripts: true,
			localResourceRoots: [
				this._extensionUri,
				vscode.Uri.joinPath(this._extensionUri, 'resource'),
				vscode.Uri.file(path.dirname(document.uri.fsPath)), // Allow access to original files
				vscode.Uri.file(os.tmpdir()) // Allow access to temp PDF files
			]
		};

		// Set initial HTML
		webviewPanel.webview.html = this._getHtmlForWebview(webviewPanel.webview);

		// Handle messages from webview
		webviewPanel.webview.onDidReceiveMessage(
			message => {
				switch (message.command) {
					case 'ready':
						this._loadPPT(document.uri.fsPath, webviewPanel);
						// Use async approach to handle file selection state after PPTX viewer is ready
						this._handleFileSelection();
						return;
					case 'retryLibreOfficeCheck':
						// Retry LibreOffice check and conversion
						this._loadPPT(document.uri.fsPath, webviewPanel);
						return;
					case 'error':
						vscode.window.showErrorMessage(`PPT Preview Error: ${message.text}`);
						return;
				}
			},
			undefined,
			[]
		);
	}

	private async _handleFileSelection() {
		// Wait for PPTX viewer to fully initialize, then ensure file is properly selected in VSCode
		setTimeout(async () => {
			// Ensure file is properly selected in VSCode
			await vscode.commands.executeCommand('workbench.files.action.focusFilesExplorer');
		}, 2000); // Add delay to ensure webview is fully loaded
	}

	private async _loadPPT(pptPath: string, webviewPanel: vscode.WebviewPanel) {
		try {
			// Check if LibreOffice is installed first
			const converter = new LibreOfficeConverter();
			const isInstalled = await converter.checkLibreOfficeInstallation();

			if (!isInstalled) {
				// Show LibreOffice installation prompt in webview
				webviewPanel.webview.postMessage({
					command: 'libreOfficeNotInstalled'
				});
				return;
			}

			// Show progress
			await vscode.window.withProgress({
				location: vscode.ProgressLocation.Notification,
				title: "Converting PowerPoint to PDF...",
				cancellable: false
			}, async (progress) => {
				// Convert PPT to PDF
				const pdfPath = await converter.convertToPDF(pptPath);

				// Load PDF in webview using PDF.js
				await this._loadPDF(pdfPath, webviewPanel);
			});
		} catch (error) {
			vscode.window.showErrorMessage(`Failed to convert PowerPoint file: ${error}`);
			webviewPanel.webview.postMessage({
				command: 'error',
				text: `Failed to convert PowerPoint file: ${error}`
			});
		}
	}

	private async _loadPDF(pdfPath: string, webviewPanel: vscode.WebviewPanel) {
		console.log('_loadPDF called with path:', pdfPath);

		if (!fs.existsSync(pdfPath)) {
			console.error('PDF file not found:', pdfPath);
			webviewPanel.webview.postMessage({
				command: 'error',
				text: 'PDF file was not created successfully.'
			});
			return;
		}

		// Convert PDF to webview URI
		const pdfUri = webviewPanel.webview.asWebviewUri(vscode.Uri.file(pdfPath));

		// Send PDF URI to webview for PDF.js to load
		webviewPanel.webview.postMessage({
			command: 'loadPDF',
			pdfUri: pdfUri.toString()
		});

		console.log('PDF URI sent to webview:', pdfUri.toString());
	}

	private _getHtmlForWebview(webview: vscode.Webview) {
		// Get PDF.js resources from extension
		const pdfResourceUri = webview.asWebviewUri(vscode.Uri.joinPath(this._extensionUri, 'resource', 'pdf'));

		// Read the viewer.html template and modify it
		const viewerHtmlPath = path.join(this._extensionUri.fsPath, 'resource', 'pdf', 'viewer.html');
		let viewerHtml = fs.readFileSync(viewerHtmlPath, 'utf8');

		// Replace the base URL placeholder
		viewerHtml = viewerHtml.replace('{{baseUrl}}', pdfResourceUri.toString());

		// Add custom script for VSCode integration
		const customScript = `
		<script>
			let pendingPdfUri = null;
			
			// Function to retry LibreOffice check
			function retryLibreOfficeCheck() {
				// Remove the prompt
				const prompt = document.getElementById('libreOfficePrompt');
				if (prompt) {
					prompt.remove();
				}
				
				// Show PDF viewer again
				const outerContainer = document.getElementById('outerContainer');
				if (outerContainer) {
					outerContainer.style.display = 'block';
				}
				
				// Notify extension to retry
				if (typeof vscode !== 'undefined') {
					vscode.postMessage({ command: 'retryLibreOfficeCheck' });
				}
			}
			
			// Function to show LibreOffice installation prompt
			function showLibreOfficePrompt() {
				// Hide PDF viewer
				const outerContainer = document.getElementById('outerContainer');
				if (outerContainer) {
					outerContainer.style.display = 'none';
				}
				
				// Create installation prompt
				const promptHtml = \`
					<div id="libreOfficePrompt" style="
						position: fixed;
						top: 0;
						left: 0;
						width: 100%;
						height: 100%;
						background: #f5f5f5;
						display: flex;
						justify-content: center;
						align-items: center;
						font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
						z-index: 10000;
					">
						<div style="
							background: white;
							padding: 40px;
							border-radius: 8px;
							box-shadow: 0 4px 20px rgba(0,0,0,0.1);
							max-width: 600px;
							text-align: center;
						">
							<div style="font-size: 48px; margin-bottom: 20px;">📄</div>
							<h2 style="color: #333; margin-bottom: 16px;">LibreOffice Required</h2>
							<p style="color: #666; margin-bottom: 24px; line-height: 1.6;">
								To preview PowerPoint files, LibreOffice needs to be installed. Please download and install it to the default system location.
							</p>
							<div style="margin-bottom: 24px;">
								<a href="https://www.libreoffice.org/download/download/" 
								   target="_blank" 
								   style="
									   display: inline-block;
									   background: #007acc;
									   color: white;
									   padding: 12px 24px;
									   text-decoration: none;
									   border-radius: 4px;
									   font-weight: 500;
									   margin: 0 8px;
								   ">Download LibreOffice</a>
								<button onclick="retryLibreOfficeCheck()" 
								        style="
									        background: #f0f0f0;
									        color: #333;
									        padding: 12px 24px;
									        border: none;
									        border-radius: 4px;
									        font-weight: 500;
									        margin: 0 8px;
									        cursor: pointer;
								        ">Retry After Installation</button>
							</div>
							<div style="font-size: 14px; color: #999;">
								<p>After installation, click "Retry After Installation" to preview your file.</p>
								<p>Supported formats: .pptx, .ppt</p>
							</div>
						</div>
					</div>
				\`;
				
				document.body.insertAdjacentHTML('beforeend', promptHtml);
			}
			
			// Function to open PDF when PDF.js is ready
			function openPdfWhenReady(pdfUri) {
				if (window.PDFViewerApplication && window.PDFViewerApplication.initializedPromise) {
					window.PDFViewerApplication.initializedPromise.then(() => {
						console.log('PDF.js initialized, opening PDF:', pdfUri);
						window.PDFViewerApplication.open(pdfUri);
					}).catch(err => {
						console.error('PDF.js initialization failed:', err);
					});
				} else {
					console.log('PDF.js not ready, waiting...');
					pendingPdfUri = pdfUri;
				}
			}
			
			// Listen for messages from extension
			window.addEventListener('message', event => {
				const message = event.data;
				switch (message.command) {
					case 'loadPDF':
						console.log('Received loadPDF command:', message.pdfUri);
						openPdfWhenReady(message.pdfUri);
						break;
					case 'libreOfficeNotInstalled':
						console.log('LibreOffice not installed, showing prompt');
						showLibreOfficePrompt();
						break;
				}
			});
			
			// Wait for PDF.js to be ready
			window.addEventListener('load', () => {
				console.log('Page loaded, checking PDF.js status');
				
				// Notify extension that webview is ready
				if (typeof vscode !== 'undefined') {
					vscode.postMessage({ command: 'ready' });
				}
				
				// Check if we have a pending PDF to load
				if (pendingPdfUri) {
					setTimeout(() => {
						openPdfWhenReady(pendingPdfUri);
					}, 1000);
				}
			});
		</script>
		`;

		// Insert the custom script before closing head tag
		viewerHtml = viewerHtml.replace('</head>', customScript + '</head>');

		return viewerHtml;
	}
}