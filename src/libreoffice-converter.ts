import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { spawn } from 'child_process';
import * as os from 'os';

export class LibreOfficeConverter {
	private getLibreOfficePath(): string {
		const platform = os.platform();

		if (platform === 'darwin') {
			// macOS
			return '/Applications/LibreOffice.app/Contents/MacOS/soffice';
		} else if (platform === 'win32') {
			// Windows - try common installation paths
			const possiblePaths = [
				'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
				'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe'
			];

			for (const possiblePath of possiblePaths) {
				if (fs.existsSync(possiblePath)) {
					return possiblePath;
				}
			}

			// If not found in common paths, try to find it in PATH
			return 'soffice.exe';
		} else {
			// Linux
			return 'libreoffice';
		}
	}

	public async checkLibreOfficeInstallation(): Promise<boolean> {
		const libreOfficePath = this.getLibreOfficePath();

		return new Promise((resolve) => {
			const process = spawn(libreOfficePath, ['--version'], { stdio: 'pipe' });

			process.on('error', () => {
				resolve(false);
			});

			process.on('close', (code) => {
				resolve(code === 0);
			});

			// Timeout after 5 seconds
			setTimeout(() => {
				process.kill();
				resolve(false);
			}, 5000);
		});
	}

	async convertToPDF(pptPath: string): Promise<string> {
		// Check if LibreOffice is installed
		const isInstalled = await this.checkLibreOfficeInstallation();
		if (!isInstalled) {
			throw new Error('LibreOffice is not installed or not found in PATH. Please install LibreOffice to use this extension.');
		}

		// Create temp directory for conversion
		const tempDir = path.join(os.tmpdir(), 'muty-pptviewer');
		if (!fs.existsSync(tempDir)) {
			fs.mkdirSync(tempDir, { recursive: true });
		}

		// Generate output PDF path
		const fileName = path.basename(pptPath, path.extname(pptPath));
		const pdfPath = path.join(tempDir, `${fileName}.pdf`);

		// Check if PDF already exists and is newer than PPT
		if (fs.existsSync(pdfPath)) {
			const pptStats = fs.statSync(pptPath);
			const pdfStats = fs.statSync(pdfPath);
			if (pdfStats.mtime > pptStats.mtime) {
				return pdfPath;
			}
		}

		return new Promise((resolve, reject) => {
			const libreOfficePath = this.getLibreOfficePath();
			const args = [
				'--headless',
				'--convert-to',
				'pdf',
				'--outdir',
				tempDir,
				pptPath
			];

			console.log(`Converting: ${libreOfficePath} ${args.join(' ')}`);

			const process = spawn(libreOfficePath, args, { stdio: 'pipe' });

			let errorOutput = '';

			process.stderr.on('data', (data) => {
				errorOutput += data.toString();
			});

			process.on('error', (error) => {
				reject(new Error(`Failed to start LibreOffice: ${error.message}`));
			});

			process.on('close', (code) => {
				if (code === 0) {
					// Check if PDF was created
					if (fs.existsSync(pdfPath)) {
						resolve(pdfPath);
					} else {
						reject(new Error('PDF file was not created. LibreOffice may have encountered an error.'));
					}
				} else {
					reject(new Error(`LibreOffice conversion failed with code ${code}. Error: ${errorOutput}`));
				}
			});

			// Timeout after 30 seconds
			setTimeout(() => {
				process.kill();
				reject(new Error('Conversion timeout. The PowerPoint file may be too large or complex.'));
			}, 30000);
		});
	}

	async convertToImages(pptPath: string): Promise<string[]> {
		// First convert to PDF
		const pdfPath = await this.convertToPDF(pptPath);

		// Then convert PDF to images using LibreOffice
		const tempDir = path.join(os.tmpdir(), 'muty-pptviewer');
		const fileName = path.basename(pptPath, path.extname(pptPath));

		return new Promise((resolve, reject) => {
			const libreOfficePath = this.getLibreOfficePath();
			const args = [
				'--headless',
				'--convert-to',
				'png',
				'--outdir',
				tempDir,
				pdfPath
			];

			console.log(`Converting PDF to images: ${libreOfficePath} ${args.join(' ')}`);

			const process = spawn(libreOfficePath, args, { stdio: 'pipe' });

			let errorOutput = '';

			process.stderr.on('data', (data) => {
				errorOutput += data.toString();
			});

			process.on('error', (error) => {
				reject(new Error(`Failed to start LibreOffice for image conversion: ${error.message}`));
			});

			process.on('close', (code) => {
				if (code === 0) {
					// Find all generated PNG files
					const imageFiles: string[] = [];
					try {
						const files = fs.readdirSync(tempDir);
						for (const file of files) {
							if (file.startsWith(fileName) && file.endsWith('.png')) {
								imageFiles.push(path.join(tempDir, file));
							}
						}

						if (imageFiles.length > 0) {
							// Sort files to ensure proper order
							imageFiles.sort();
							resolve(imageFiles);
						} else {
							reject(new Error('No PNG files were created from PDF.'));
						}
					} catch (error) {
						reject(new Error(`Error reading converted images: ${error}`));
					}
				} else {
					reject(new Error(`LibreOffice image conversion failed with code ${code}. Error: ${errorOutput}`));
				}
			});

			// Timeout after 30 seconds
			setTimeout(() => {
				process.kill();
				reject(new Error('Image conversion timeout.'));
			}, 30000);
		});
	}
}
