import fse from 'fs-extra';
import path from 'path';
const topDir = import.meta.dirname;
fse.emptyDirSync(path.join(topDir, 'public', 'tinymce'));
fse.copySync(path.join(topDir, 'node_modules', 'tinymce'), path.join(topDir, 'public', 'tinymce'), { overwrite: true });

if (process.env.VERCEL === '1') {
	const envVercelPath = path.join(topDir, '.env.vercel');
	const envPath = path.join(topDir, '.env');
	if (fse.existsSync(envVercelPath)) {
		fse.copySync(envVercelPath, envPath, { overwrite: true });
		console.log('Copied .env.vercel to .env for Vercel deployment');
	}
}

